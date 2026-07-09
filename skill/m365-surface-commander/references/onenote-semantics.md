---
title: OneNote Semantics
kind: reference
skill: m365-surface-commander
surface: onenote
topics: [page, paragraph, note, append-page]
load_when: The active surface is OneNote or a plan mentions pages, notebook notes, or append-page.
---

# OneNote semantics

Load this when the active surface is OneNote. Cross-surface table:
[capability-map.md](capability-map.md). OneNote is **web-only** (the add-in runs only in OneNote on
the web) and has **no event API**, so there is no live `watch` — you act on the snapshot you're given.

**Reading:** `read`/`outline`/`search` over the page/section. Synthesized content lands as HTML in the
supported subset; untrusted text is HTML-escaped.

**Progressive disclosure:** build slash commands from the live OneNote capability scan for the active
web page/selection. Prefer `read` / `list` / `inspect` / `properties` / `open` before writes: read the
current page, list notebooks/sections/pages when available, inspect selection/paragraph anchors and HTML
support, fetch page properties, then open the intended page/section. If the active target, insertion
point, image source, or object capability is missing, fail closed and do not synthesize a write target.

**Core verb:** `page "Title" "body"` appends a synthesized page (with inline citation tags when
sources are present).

**Specialized (`/`) when live-capability advertised:** `/set-page-title` changes the active page title;
`/add-outline` inserts an outline onto the active page; `/insert-table` and `/insert-image` insert only
after the destination and source are resolved; `/add-note-tag` applies supported built-in tags (To-Do /
Important / Question) to a resolved paragraph/selection. Also available when gated: `/append-rich-text`,
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
