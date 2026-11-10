# Project rules

- Work only inside this repository (`D:\dima`). Never read or write files outside
  it: no home directory, AppData, temp, other drives, or `..` escapes. If a
  request needs an outside path, stop and reply exactly:
  «نمی‌تونم — فقط داخل پوشهٔ پروژه کار می‌کنم.»
- The Telegram bot (`tools/telegram_bot.py`) runs headless agents under the same
  rule, and `opencode.json` denies `external_directory` for tool access.
- Every change must be committed and pushed to `origin/main` before the reply ends.
  No exceptions, no "later", no leaving work uncommitted.
- Commit messages: short, imperative, English (e.g. `Fix blockchain slide wording`).
- The published site is `presenter-guide/`. Anything the site links to must live inside that folder.
- Deck typography is B Nazanin (body) + B Titr (titles), normal spaces instead of ZWNJ.
- The deck is generated from `build_deck.py`. Never hand-edit `FL_IoT_Privacy_FA.pptx` in
  PowerPoint and save over it: the next build overwrites those edits. Any manual tweak must be
  ported into `build_deck.py` first, then rebuilt and re-exported to `presenter-guide/slides/`.
