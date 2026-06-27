# OneNote semantics

Load this when the active surface is OneNote. Cross-surface table:
[capability-map.md](capability-map.md). OneNote is **web-only** (the add-in runs only in OneNote on
the web) and has **no event API**, so there is no live `watch` — you act on the snapshot you're given.

**Reading:** `read`/`outline`/`search` over the page/section. Synthesized content lands as HTML in the
supported subset; untrusted text is HTML-escaped.

**Core verb:** `page "Title" "body"` appends a synthesized page (with inline citation tags when
sources are present).

**Specialized (`/`):** `/add-outline` (onto the active page), `/insert-table`, `/insert-image`,
`/append-rich-text`, `/set-page-title`, `/add-note-tag` (To-Do / Important / Question),
`/create-section`, `/insert-link-at-cursor`.

**Gotchas — reversibility is weak here.** Most OneNote objects (`Outline`, `Image`, `Table`, `Page`,
`Section`) have **no `.delete()`** in the API; the only programmatic deletes are a paragraph and a
page-content region. So:

- `/set-page-title` is the one **fully reversible** write (restore the prior title) — prefer it.
- `/add-outline` / `/insert-table` / `/insert-image` undo by deleting their parent region/paragraph.
- `append-page`, `/create-section`, and cursor-relative inserts have **no in-API undo** — they are
  gated, and true deletion needs the Graph estate path. Treat them as not cleanly reversible.

Targeted replace and real page/section deletion exist **only on Graph** (`/graph-patch-page`,
`/graph-create-section`) — Plane B, needs estate auth.
