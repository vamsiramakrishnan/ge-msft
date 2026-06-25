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
**Specialized (`/`):** `/insert-table`, `/insert-image`, `/insert-html`, `/apply-style`,
`/apply-list`, `/insert-field` (TOC / cross-ref / page number), `/insert-bookmark`,
`/insert-content-control`, `/fill-content-control`, `/insert-break`, `/set-header-footer`,
`/insert-note` (foot/endnote), `/find-replace`, `/set-doc-properties`, `/insert-paragraph`.

**Gotchas:** Word has **no chart and no equation API** — embed those via OOXML (`/insert-ooxml`) or as
an image. `/set-page-layout` and accept/reject-*all* revisions are desktop-only (degrade on the web).
`/resolve-revisions` is **irreversible** — it discards the alternative text; it warns hard.
