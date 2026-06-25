# PowerPoint semantics

Load this when the active surface is PowerPoint. Cross-surface table:
[capability-map.md](capability-map.md).

**Reading:** `read`/`outline` give the slides and their shapes. Write by slide index or by an existing
shape's id/name.

**Core verb:** `slide "Title" "bullet" …` inserts a slide (bullets can be a table expression whose
rows become bullets).

**Specialized (`/`) — the authoring surface:** `/add-shape` (text box / geometric / line),
`/add-table-slide` (a native table with seeded values), `/set-shape-text` (edit an existing shape),
`/format-shape` (fill/line/font/geometry), `/delete-slide`, `/move-slide`, `/duplicate-slide`,
`/apply-slide-layout`, `/insert-hyperlink`, `/insert-image`.

**Reversibility:** shapes/tables/slides are name/id-anchored and deletable, so most writes record a
clean inverse (delete the minted object; a deleted slide is snapshotted for restore).

**Gotchas:** PowerPoint's JS API is the narrowest — there is **no chart-from-data**, **no SmartArt**,
**no speaker-notes writer** (`set-speaker-notes` stays unadvertised), and **no loose-OOXML insert**.
`/insert-image` goes through the Common API and returns no shape handle, so it has **no clean undo** —
it is gated hard. Prefer the modeled shape/table kinds where they exist.
