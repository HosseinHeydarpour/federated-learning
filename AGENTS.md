# Project rules

- Every change must be committed and pushed to `origin/main` before the reply ends.
  No exceptions, no "later", no leaving work uncommitted.
- Commit messages: short, imperative, English (e.g. `Fix blockchain slide wording`).
- The published site is `presenter-guide/`. Anything the site links to must live inside that folder.
- Deck typography is B Nazanin (body) + B Titr (titles), normal spaces instead of ZWNJ.
- The deck is generated from `build_deck.py`. Never hand-edit `FL_IoT_Privacy_FA.pptx` in
  PowerPoint and save over it: the next build overwrites those edits. Any manual tweak must be
  ported into `build_deck.py` first, then rebuilt and re-exported to `presenter-guide/slides/`.
