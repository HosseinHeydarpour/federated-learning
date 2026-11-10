# -*- coding: utf-8 -*-
"""Telegram bridge for the presenter guide.

Presenters send a change request; the bot hands it to a headless OpenCode run in
this repo, which edits, tests, commits and pushes per AGENTS.md, then replies
with the summary.

Run:  python tools/telegram_bot.py
Config: tools/.env (TELEGRAM_BOT_TOKEN=...) and tools/allowlist.txt
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
ENV_FILE = TOOLS / ".env"
ALLOW_FILE = TOOLS / "allowlist.txt"
CLI_GLOB = Path(os.environ.get("APPDATA", "")) / "ai.opencode.desktop" / "cli"
MAX_MSG = 3800
AGENT_TIMEOUT = 1800
DEFAULT_MODEL = "opencode-go/deepseek-v4.1-flash"

HELP = """ربات تغییرات راهنمای ارائه.

هر پیام معمولی را به‌عنوان درخواست تغییر می‌فرستم؛ ربات خودش فایل‌ها را عوض می‌کند، تست می‌گیرد و روی گیت‌هاب پوش می‌کند.

دستورها:
/status — وضعیت مخزن و آخرین کامیت‌ها
/ask متن — همان درخواست تغییر (اگر پیام با / شروع شود)
/id — شناسهٔ تلگرام تو (برای اضافه‌شدن به لیست مجاز)
/help — همین راهنما

نمونه: «عنوان اسلاید ۹ را با متن اسلاید یکی کن و پوش کن.»"""


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


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


def send(chat_id, text):
    for start in range(0, len(text), MAX_MSG):
        api("sendMessage", {"chat_id": chat_id, "text": text[start:start + MAX_MSG],
                            "disable_web_page_preview": True})


def allowed_ids():
    if not ALLOW_FILE.exists():
        return set()
    ids = set()
    for line in ALLOW_FILE.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if line.lstrip("-").isdigit():
            ids.add(int(line))
    return ids


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


def run_agent(request, sender):
    cli = find_cli()
    if not cli:
        return "خطا: opencode-cli پیدا نشد. OPENCODE_CLI را در tools/.env بگذار."
    prompt = (
        f"درخواست یکی از ارائه‌دهنده‌ها از تلگرام ({sender}):\n{request}\n\n"
        "قواعد: فقط داخل همین مخزن کار کن و هیچ فایل بیرون از آن را تغییر نده. "
        "اگر درخواست مبهم است حدس نزن؛ در پایان یک سؤال کوتاه بپرس. "
        "تغییر را با اسکریپت‌های پروژه بساز و تست کن. "
        "طبق AGENTS.md همه‌چیز را کامیت و روی origin/main پوش کن. "
        "پاسخ نهایی حداکثر ۵ خط فارسی: چه عوض شد، هش کامیت، و اگر سؤالی هست."
    )
    cmd = [cli, "run", "--auto", "--title", f"تلگرام: {sender}"]
    model = os.environ.get("OPENCODE_MODEL", "").strip() or DEFAULT_MODEL
    cmd += ["--model", model]
    cmd.append(prompt)
    log("agent start:", request[:60].replace("\n", " "))
    try:
        proc = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=AGENT_TIMEOUT)
    except subprocess.TimeoutExpired:
        return "زمان اجرا تمام شد؛ تغییرات نیمه‌کاره ممکن است مانده باشد. /status را بزن."
    out = strip_ansi((proc.stdout or "").strip())
    err = strip_ansi((proc.stderr or "").strip())
    out = "\n".join(line for line in out.splitlines() if not line.startswith("> ")).strip()
    if proc.returncode != 0:
        return f"اجرا با کد {proc.returncode} تمام شد.\n{out[-1500:]}\n{err[-800:]}".strip()
    return out or "انجام شد، ولی خروجی متنی نبود. /status را بزن."


def git_status():
    def git(*args):
        p = subprocess.run(["git", *args], cwd=str(ROOT), capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
        return (p.stdout or p.stderr).strip()
    return git("status", "-sb") + "\n\n" + git("log", "--oneline", "-3")


def handle(msg):
    chat_id = msg["chat"]["id"]
    user = msg.get("from", {})
    uid = user.get("id")
    name = user.get("first_name") or user.get("username") or str(uid)
    text = (msg.get("text") or "").strip()
    if not text:
        return

    if uid not in allowed_ids():
        send(chat_id, f"این ربات خصوصی است.\nشناسهٔ تلگرام تو: {uid}\n"
                      "این عدد را به مسئول پروژه بده تا اضافه شود.")
        log("denied", uid, name)
        return

    if text in ("/start", "/help"):
        send(chat_id, HELP)
    elif text == "/id":
        send(chat_id, f"شناسهٔ تلگرام: {uid}\nشناسهٔ چت: {chat_id}")
    elif text == "/status":
        send(chat_id, git_status())
    elif text.startswith("/ask"):
        request = text[4:].strip()
        if not request:
            send(chat_id, "بعد از /ask متن درخواست را بنویس.")
            return
        send(chat_id, "⏳ در حال اعمال؛ ممکن است چند دقیقه طول بکشد.")
        send(chat_id, run_agent(request, name))
    elif text.startswith("/"):
        send(chat_id, "دستور ناشناخته. /help")
    else:
        send(chat_id, "⏳ در حال اعمال؛ ممکن است چند دقیقه طول بکشد.")
        send(chat_id, run_agent(text, name))


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
    print("allowlist:", sorted(allowed_ids()) or "empty (everyone will be refused)")
    print("cli:", find_cli() or "not found")
    print("git:\n" + git_status())
    return 0 if (os.environ.get("TELEGRAM_BOT_TOKEN") or "--allow-missing-token" in sys.argv) else 1


def main():
    load_env()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if "--selftest" in sys.argv:
        return selftest()
    if not allowed_ids():
        log("warning: tools/allowlist.txt is empty; the bot refuses everyone and prints their id")
    offset = load_offset()
    log("bot started; repo:", ROOT)
    while True:
        try:
            updates = api("getUpdates", {"offset": offset, "timeout": 30,
                                         "allowed_updates": ["message"]}, timeout=60)
            for upd in updates:
                offset = upd["update_id"] + 1
                save_offset(offset)
                if "message" in upd:
                    handle(upd["message"])
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
