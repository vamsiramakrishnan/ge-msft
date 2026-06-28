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

**Gotchas:** Word has **no chart and no equation API** in the current bridge. Other modeled Word
kinds (`/insert-table`, `/insert-image`, `/apply-style`, `/set-header-footer`, revision resolution,
and document-property writes) are not active unless they appear in the live grammar/help for the
turn.
