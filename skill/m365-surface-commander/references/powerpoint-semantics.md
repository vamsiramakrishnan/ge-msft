---
title: PowerPoint Semantics
kind: reference
skill: m365-surface-commander
surface: powerpoint
topics: [slides, shapes, insert-slide, navigation]
load_when: The active surface is PowerPoint or a handoff targets a PowerPoint deck.
---

# PowerPoint semantics

Load this when the active surface is PowerPoint. Cross-surface table:
[capability-map.md](capability-map.md).

**Reading:** `read`/`outline` give the slides and their shapes. Write by slide index or by an existing
shape ref. Use `list shape`, `properties <ref>`, `open <ref>`, and `inspect <ref>` to target one
shape; do not parse a shape id out of slide text.

**Core verbs:**

- `slide "Title" "bullet" …` inserts a slide (bullets can be a table expression whose rows become
  bullets).
- `shape <pp:shape:slideId:shapeId> "text"` replaces text in one existing shape/text box. Discover
  the target with `list shape`, `properties <ref>`, `open <ref>`, and `inspect <ref>` first.

**Generated deck artifact path:** for multi-slide generated decks, the host app may compile a
bounded DeckSpec/HTML-derived preview into one base64 `.pptx` artifact and invoke `insert-slide`
with `params.deck`. Do **not** emit raw base64 PPTX in a normal command block; use `slide` for
ordinary one-slide creation unless the runtime has already staged a compiled deck artifact.

**Specialized (`/`) today:** use only what the live grammar advertises. PowerPoint has the bare
`slide` and `shape` verbs above; the generated-deck import is an internal client-staged
`insert-slide` artifact path, not a conversational shortcut. Modeled future kinds such as
`/add-shape`, `/add-table-slide`, `/format-shape`, `/delete-slide`, `/move-slide`,
`/duplicate-slide`, `/apply-slide-layout`, `/insert-hyperlink`, and `/insert-image` must not be used
until they appear in the live grammar/help for the turn.

**Reversibility:** inserted slides and exact shape text replacements record bridge-level inverses.

**Gotchas:** PowerPoint's JS API is the narrowest — there is **no chart-from-data**, **no SmartArt**,
**no speaker-notes writer** (`set-speaker-notes` stays unadvertised), and **no loose-OOXML insert**.
Image/table/shape creation beyond those two writes stays modeled until the bridge has a tested host
write path and inverse.
