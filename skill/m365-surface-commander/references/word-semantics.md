---
title: Word Semantics
kind: reference
skill: m365-surface-commander
surface: word
topics: [tracked-changes, comments, anchors, paragraphs]
load_when: The active surface is Word or a plan mentions redlines, comments, paragraphs, or text anchors.
---

# Word semantics

Load this when the active surface is Word. Cross-surface table: [capability-map.md](capability-map.md).

**Reading:** `read` returns the whole document (or a section); `search <text>` locates passages. You
**cannot anchor on text you have not read** — read/search first.

**Writing is content-anchored.** `suggest "exact old" => "new"` lands a **tracked change** anchored on
the exact existing text; `comment "exact text" "note"` flags it. The anchor is re-resolved at
apply-time: if the text has drifted, the change **degrades to a panel item** rather than landing a
broken annotation. Never anchor on approximate or guessed text.

**Tracked changes are the default reversibility model** — agent edits land as reviewable revisions,
not silent rewrites. A correction the source supports → `suggest`; an unverifiable claim → `comment`,
never a silent change.

**Core verbs:** `suggest`, `comment`, `reply`.

**Progressive disclosure:** use only the slash commands advertised by the live grammar/help for the
turn. Advanced Office.js-backed verbs can percolate into CLI, skills, and UI as they become
capability-gated; absence from the live grammar means **do not call them**. Fail closed: read/search
or `list`/`open`/`properties`/`inspect` first, and if the target, style, id, or command is missing,
fall back to `suggest`/`comment` or a panel item instead of inventing a write.

**Specialized (`/`) currently advertised by the Word bridge:**

- `/insert-text text="..." [match="exact anchor"] [contextHint="..."]` inserts plain text at the
  current selection or after an exact anchor. It is direct and currently irreversible, so use
  `suggest` for normal rewrites.
- `/replace-selection text="..."` replaces the current selection. The bridge captures prior selected
  text as a restore inverse. Never retarget to similar text if the selection is gone.
- `/insert-ooxml ooxml="<w:p/>" [match="exact anchor"] [contextHint="..."]` inserts a small screened
  OOXML fragment. Use only when plain text or tracked changes cannot express the requested Word
  structure.
- `/fill-content-control id=<contentControlId> text="..."` fills a known content control id. Find the
  id from host context/properties; never infer it from visible text.

**Advanced (`/`) when live-advertised:**

- `/apply-style` applies a known Word style to an inspected range, paragraph, or selection. Discover
  style names and target refs first; do not free-type a style that the document does not expose.
- `/insert-table` creates a Word table from structured rows at the selection or after an exact anchor.
  Prefer it over OOXML for normal tables; verify the anchor/selection before writing.
- `/insert-image` inserts an image at the selection, anchor, or inspected content range. Use only with
  a staged/approved image payload or URL accepted by the bridge, and keep alt text explicit.
- `/insert-content-control` creates a bound or plain content control only when the requested tag/title
  and insertion target are known. Subsequent fills must use inspected content-control ids.
- Paragraph/list/style flow should be explicit: read the surrounding paragraphs, inspect list/style
  state, then apply paragraph formatting, list creation/continuation, or style changes in that order.

**Gotchas:** Word has **no chart and no equation API** in the current bridge. Other modeled Word
kinds (`/set-header-footer`, revision resolution, and document-property writes) are not active
unless they appear in the live grammar/help for the turn.
