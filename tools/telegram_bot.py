# -*- coding: utf-8 -*-
"""Telegram bridge for the presenter guide.

Presenters send a change request. It waits for the project manager's approval
unless the sender is an admin. Approved requests go to a headless OpenCode run
in this repo, which edits, tests, commits and pushes per AGENTS.md, then the
result is sent back as clean Telegram HTML.

Run:  python tools/telegram_bot.py
Config: tools/.env (token), tools/allowlist.txt, tools/admins.txt
"""
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
STATE = TOOLS / "telegram-state.json"
PENDING = TOOLS / "pending.json"
LOG_FILE = TOOLS / "bot.log"
ENV_FILE = TOOLS / ".env"
ALLOW_FILE = TOOLS / "allowlist.txt"
ADMIN_FILE = TOOLS / "admins.txt"
CLI_GLOB = Path(os.environ.get("APPDATA", "")) / "ai.opencode.desktop" / "cli"
MAX_MSG = 3800
AGENT_TIMEOUT = 1800
DEFAULT_MODEL = "opencode-go/deepseek-v4.1-flash"

REFUSAL = ("نمی‌تونم — فقط داخل پوشهٔ پروژه کار می‌کنم.\n"
           "این درخواست به فایل یا پوشه‌ای بیرون از مخزن نیاز دارد.")

OUTSIDE_PATH_RE = re.compile(
    r"[A-Za-z]:[\\/][^\s\"'<>|]*"          # C:\...
    r"|\\\\[^\\\s]+"                        # \\server\share
    r"|(?<![\w])~[\\/]"                     # ~/
    r"|(?<![\w])\.\.[\\/]"                  # ..\ or ../
    r"|%(?:appdata|userprofile|temp|programdata|windir|systemroot)%"
    r"|(?:^|\s)/(?:etc|home|usr|var|root|opt|bin|tmp)/",
    re.I,
)
OUTSIDE_WORDS = ("دسکتاپ", "دانلودها", "خارج از پوشه", "خارج از پروژه",
                 "خارج از ریپو", "کل درایو", "my documents", "desktop", "downloads",
                 "مشخصات سیستم", "اطلاعات سیستم", "systeminfo", "system info")
PROJECT_WORDS = ("پروژه", "سایت", "اسلاید", "راهنما", "ریپو", "ریورت", "حذف",
                 "پادکست", "گیت", "پوش")


def outside_request(text):
    """True when the request names a path outside the repository."""
    low = text.lower()
    if any(word in low for word in OUTSIDE_WORDS) and not any(word in low for word in PROJECT_WORDS):
        return True
    for raw in OUTSIDE_PATH_RE.findall(text):
        if not raw or raw.startswith(("..", "~", "%", "\\\\", "/")):
            return True
        try:
            target = Path(raw.strip("\"'")).resolve()
        except OSError:
            return True
        if target != ROOT and ROOT not in target.parents:
            return True
    return False

HELP = """<b>ربات تغییرات راهنمای ارائه</b>

هر پیام معمولی را به‌عنوان درخواست تغییر می‌فرستم. درخواست اعضای عادی اول برای تایید به مدیر پروژه می‌رود.

دستورها:
/status — وضعیت مخزن و آخرین کامیت‌ها
/ask متن — همان درخواست تغییر
/id — شناسهٔ تلگرام تو
/help — همین راهنما

اگر مدیر پروژه هستی:
/pending — درخواست‌های در انتظار
/approve A1 · /reject A1 — تایید یا رد با شناسه

فقط داخل پوشهٔ پروژه کار می‌کنم؛ درخواست بیرونی رد می‌شود."""


def log(*a):
    line = time.strftime("%H:%M:%S") + " " + " ".join(str(x) for x in a)
    print(line, flush=True)
    try:
        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def load_env():
    if not ENV_FILE.exists():
        return
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def api(method, payload=None, timeout=40):
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
    if not token:
        raise SystemExit("TELEGRAM_BOT_TOKEN is not set (tools/.env)")
    data = json.dumps(payload or {}).encode("utf-8")
    req = urllib.request.Request(
        f"https://api.telegram.org/bot{token}/{method}",
        data=data,
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as res:
        body = json.loads(res.read().decode("utf-8"))
    if not body.get("ok"):
        raise RuntimeError(body.get("description", "telegram error"))
    return body["result"]


def esc(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def md_to_html(text):
    out = esc(text)
    out = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", out)
    out = re.sub(r"`([^`\n]+)`", r"<code>\1</code>", out)
    return out


def send(chat_id, text, markup=None):
    payload = {"chat_id": chat_id, "text": str(text)[:MAX_MSG], "parse_mode": "HTML",
               "disable_web_page_preview": True}
    if markup:
        payload["reply_markup"] = markup
    api("sendMessage", payload)


def edit(chat_id, message_id, text, markup=None):
    payload = {"chat_id": chat_id, "message_id": message_id,
               "text": str(text)[:MAX_MSG], "parse_mode": "HTML",
               "disable_web_page_preview": True}
    if markup:
        payload["reply_markup"] = markup
    try:
        api("editMessageText", payload)
    except RuntimeError as e:
        if "not modified" not in str(e):
            log("edit failed:", e)


def allowed_ids(path):
    if not path.exists():
        return set()
    ids = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if line.lstrip("-").isdigit():
            ids.add(int(line))
    return ids


def load_pending():
    try:
        data = json.loads(PENDING.read_text(encoding="utf-8"))
        data.setdefault("seq", 0)
        data.setdefault("items", {})
        return data
    except Exception:
        return {"seq": 0, "items": {}}


def save_pending(data):
    try:
        PENDING.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError as e:
        log("pending save failed:", e)


def pending_add(uid, name, chat, text):
    data = load_pending()
    data["seq"] += 1
    rid = f"A{data['seq']}"
    data["items"][rid] = {"uid": uid, "name": name, "chat": chat, "text": text,
                          "ts": int(time.time())}
    save_pending(data)
    return rid


def pending_pop(rid):
    data = load_pending()
    item = data["items"].pop(rid, None)
    save_pending(data)
    return item


def find_cli():
    env = os.environ.get("OPENCODE_CLI", "").strip()
    if env and Path(env).exists():
        return env
    found = sorted(
        CLI_GLOB.glob("*/opencode-cli.exe"),
        key=lambda p: [int(x) for x in re.findall(r"\d+", p.parent.name)],
        reverse=True,
    )
    return str(found[0]) if found else ""


def strip_ansi(text):
    return re.sub(r"\x1b\[[0-9;]*m", "", text)


def parse_texts(stdout):
    """Collect assistant text parts from `opencode run --format json` output."""
    texts = []
    for line in strip_ansi(stdout or "").splitlines():
        line = line.strip()
        if not line.startswith("{"):
            continue
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        if event.get("type") == "text":
            part = event.get("part") or {}
            value = (part.get("text") or "").strip()
            if value:
                texts.append(value)
    return texts


def run_agent(request, sender):
    cli = find_cli()
    if not cli:
        return "خطا: opencode-cli پیدا نشد. OPENCODE_CLI را در tools/.env بگذار."
    prompt = (
        f"درخواست یکی از ارائه‌دهنده‌ها از تلگرام ({sender}):\n{request}\n\n"
        "مرز دسترسی: فقط داخل ریشهٔ همین مخزن کار کن. هیچ مسیر مطلقی بیرون از آن را "
        "باز نکن، حتی برای خواندن. اگر درخواست به فایل یا پوشه‌ای بیرون از مخزن نیاز داشت، "
        "هیچ کاری نکن و پاسخ را با این جمله شروع کن: «نمی‌تونم — فقط داخل پوشهٔ پروژه کار می‌کنم.»\n\n"
        "قواعد: فقط داخل همین مخزن کار کن و هیچ فایل بیرون از آن را تغییر نده. "
        "درخواست اطلاعات سیستم یا فایل‌های شخصی کاربر را هم رد کن. "
        "فقط فایل‌هایی را کامیت کن که خودت در همین اجرا تغییر داده‌ای؛ فایل‌های نامرتبط یا "
        "تغییرات نیمه‌کاره را add/commit/push نکن. "
        "اگر درخواست مبهم است حدس نزن؛ در پایان یک سؤال کوتاه بپرس. "
        "تغییر را با اسکریپت‌های پروژه بساز و تست کن. "
        "طبق AGENTS.md همه‌چیز را کامیت و روی origin/main پوش کن. "
        "پاسخ نهایی حداکثر ۵ خط فارسی و بدون فهرست بلند: چه عوض شد، هش کامیت، و اگر سؤالی هست."
    )
    cmd = [cli, "run", "--auto", "--format", "json", "--title", f"تلگرام: {sender}"]
    model = os.environ.get("OPENCODE_MODEL", "").strip() or DEFAULT_MODEL
    cmd += ["--model", model, prompt]
    log("agent start:", request[:60].replace("\n", " "))
    try:
        proc = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=AGENT_TIMEOUT)
    except subprocess.TimeoutExpired:
        return "زمان اجرا تمام شد؛ تغییرات نیمه‌کاره ممکن است مانده باشد. /status را بزن."

    texts = parse_texts(proc.stdout)
    body = texts[-1] if texts else ""
    if not body:
        err = strip_ansi((proc.stderr or "").strip())[-500:]
        return f"اجرا با کد {proc.returncode} تمام شد و متنی برنگشت.\n{err}".strip()
    return body


def git_status():
    def git(*args):
        p = subprocess.run(["git", *args], cwd=str(ROOT), capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
        return (p.stdout or p.stderr).strip()
    return "<b>وضعیت مخزن</b>\n<pre>" + esc(git("status", "-sb")) + "</pre>\n<b>آخرین کامیت‌ها</b>\n<pre>" + esc(git("log", "--oneline", "-3")) + "</pre>"


def result_message(text):
    body = text.strip()
    if len(body) > 3300:
        body = body[:3300].rstrip() + "…"
    return "🤖 <b>نتیجه</b>\n\n" + md_to_html(body)


def request_change(chat_id, uid, name, text):
    if outside_request(text):
        log("refused outside:", text[:80].replace("\n", " "))
        send(chat_id, esc(REFUSAL))
        return
    if uid in allowed_ids(ADMIN_FILE):
        send(chat_id, "⏳ در حال اعمال؛ ممکن است چند دقیقه طول بکشد.")
        send(chat_id, result_message(run_agent(text, name)))
        return

    rid = pending_add(uid, name, chat_id, text)
    send(chat_id, f"⏳ درخواستت ثبت شد (<b>{rid}</b>) و برای تایید به مدیر پروژه رفت.")
    for admin in sorted(allowed_ids(ADMIN_FILE)):
        markup = {"inline_keyboard": [[
            {"text": "✅ تایید", "callback_data": f"ap:{rid}"},
            {"text": "❌ رد", "callback_data": f"rj:{rid}"},
        ]]}
        send(admin, f"📥 <b>درخواست {rid}</b>\nاز: {esc(name)}\n\n{esc(text)}", markup)


def handle(msg):
    chat_id = msg["chat"]["id"]
    user = msg.get("from", {})
    uid = user.get("id")
    name = user.get("first_name") or user.get("username") or str(uid)
    text = (msg.get("text") or "").strip()
    if not text:
        return

    if uid not in allowed_ids(ALLOW_FILE):
        send(chat_id, f"این ربات خصوصی است.\nشناسهٔ تلگرام تو: {uid}\n"
                      "این عدد را به مسئول پروژه بده تا اضافه شود.")
        log("denied", uid, name)
        return

    is_admin = uid in allowed_ids(ADMIN_FILE)

    if text in ("/start", "/help"):
        send(chat_id, HELP)
    elif text == "/id":
        send(chat_id, f"شناسهٔ تلگرام: {uid}\nشناسهٔ چت: {chat_id}\n"
                      f"نقش: {'مدیر پروژه' if is_admin else 'ارائه‌دهنده'}")
    elif text == "/status":
        send(chat_id, git_status())
    elif text == "/pending":
        if not is_admin:
            send(chat_id, "فقط مدیر پروژه.")
            return
        items = load_pending()["items"]
        if not items:
            send(chat_id, "صف خالی است.")
            return
        lines = [f"<b>در انتظار تایید</b> ({len(items)})"]
        for rid, item in list(items.items())[-10:]:
            lines.append(f"<b>{rid}</b> · {esc(item['name'])}\n{esc(item['text'][:120])}")
        send(chat_id, "\n\n".join(lines))
    elif text.startswith(("/approve", "/reject")):
        if not is_admin:
            send(chat_id, "فقط مدیر پروژه.")
            return
        cmd, _, rid = text.partition(" ")
        rid = rid.strip()
        item = pending_pop(rid)
        if not item:
            send(chat_id, "این شناسه در صف نیست.")
            return
        if cmd == "/reject":
            send(chat_id, f"❌ <b>{rid}</b> رد شد.")
            send(item["chat"], f"❌ درخواستت (<b>{rid}</b>) رد شد.")
            return
        send(chat_id, f"✅ <b>{rid}</b> تایید شد؛ در حال اجرا…")
        send(item["chat"], f"✅ درخواستت (<b>{rid}</b>) تایید شد؛ در حال اجرا…")
        result = result_message(run_agent(item["text"], item["name"]))
        send(item["chat"], result)
        send(chat_id, result)
    elif text.startswith("/ask"):
        request = text[4:].strip()
        if not request:
            send(chat_id, "بعد از /ask متن درخواست را بنویس.")
            return
        request_change(chat_id, uid, name, request)
    elif text.startswith("/"):
        send(chat_id, "دستور ناشناخته. /help")
    else:
        request_change(chat_id, uid, name, text)


def handle_callback(cq):
    uid = cq.get("from", {}).get("id")
    api("answerCallbackQuery", {"callback_query_id": cq["id"]})
    if uid not in allowed_ids(ADMIN_FILE):
        return
    msg = cq.get("message") or {}
    chat_id = (msg.get("chat") or {}).get("id")
    message_id = msg.get("message_id")
    action, _, rid = (cq.get("data") or "").partition(":")
    if action not in ("ap", "rj"):
        return
    item = pending_pop(rid)
    if not item:
        if chat_id and message_id:
            edit(chat_id, message_id, f"این درخواست (<b>{rid}</b>) دیگر در صف نیست.")
        return
    if action == "rj":
        if chat_id and message_id:
            edit(chat_id, message_id, f"❌ <b>{rid}</b> رد شد.\nاز: {esc(item['name'])}")
        send(item["chat"], f"❌ درخواستت (<b>{rid}</b>) رد شد.")
        return
    if chat_id and message_id:
        edit(chat_id, message_id, f"✅ <b>{rid}</b> تایید شد؛ در حال اجرا…\nاز: {esc(item['name'])}")
    send(item["chat"], f"✅ درخواستت (<b>{rid}</b>) تایید شد؛ در حال اجرا…")
    result = result_message(run_agent(item["text"], item["name"]))
    send(item["chat"], result)
    if chat_id:
        send(chat_id, result)


def load_offset():
    try:
        return int(json.loads(STATE.read_text(encoding="utf-8")).get("offset", 0))
    except Exception:
        return 0


def save_offset(offset):
    try:
        STATE.write_text(json.dumps({"offset": offset}), encoding="utf-8")
    except Exception as e:
        log("state save failed:", e)


def selftest():
    load_env()
    print("token:", "ok" if os.environ.get("TELEGRAM_BOT_TOKEN") else "MISSING (tools/.env)")
    print("allowlist:", sorted(allowed_ids(ALLOW_FILE)) or "empty (everyone refused)")
    print("admins:", sorted(allowed_ids(ADMIN_FILE)) or "empty (nobody can approve)")
    print("pending:", len(load_pending()["items"]))
    print("cli:", find_cli() or "not found")
    print("git:\n" + strip_ansi(git_status()).replace("<pre>", "").replace("</pre>", ""))
    return 0 if os.environ.get("TELEGRAM_BOT_TOKEN") else 1


def main():
    load_env()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if "--selftest" in sys.argv:
        return selftest()
    if not allowed_ids(ALLOW_FILE):
        log("warning: tools/allowlist.txt is empty; everyone is refused")
    if not allowed_ids(ADMIN_FILE):
        log("warning: tools/admins.txt is empty; no request can be approved")
    offset = load_offset()
    log("bot started; repo:", ROOT)
    while True:
        try:
            updates = api("getUpdates", {"offset": offset, "timeout": 30,
                                         "allowed_updates": ["message", "callback_query"]},
                          timeout=60)
            for upd in updates:
                offset = upd["update_id"] + 1
                save_offset(offset)
                if "message" in upd:
                    handle(upd["message"])
                elif "callback_query" in upd:
                    try:
                        handle_callback(upd["callback_query"])
                    except Exception as e:
                        log("callback error", repr(e))
        except urllib.error.HTTPError as e:
            log("http", e.code, e.read()[:200])
            time.sleep(5)
        except urllib.error.URLError as e:
            log("network", e.reason)
            time.sleep(10)
        except KeyboardInterrupt:
            log("stopped")
            return 0
        except Exception as e:
            log("error", repr(e))
            time.sleep(5)


if __name__ == "__main__":
    sys.exit(main())
