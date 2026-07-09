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

**Progressive disclosure:** specialized PowerPoint writes are live capability-gated. Use only slash
commands advertised by the grammar/help for the turn, after `list`/`properties`/`open`/`inspect`
confirms the slide, shape, layout, or staged asset. Fail closed: if the command, ref, layout, or
payload is absent, stop at a readable plan/comment instead of fabricating ids or silently degrading
to an unsupported Office.js path.

**Core verbs:**

- `slide "Title" "bullet" …` inserts a slide (bullets can be a table expression whose rows become
  bullets).
- `shape <pp:shape:slideId:shapeId> "text"` replaces text in one existing shape/text box. Discover
  the target with `list shape`, `properties <ref>`, `open <ref>`, and `inspect <ref>` first.

**Generated deck artifact path:** for multi-slide generated decks, the host app may compile a
bounded DeckSpec/HTML-derived preview into one base64 `.pptx` artifact and invoke `/insert-slide`
with a `deckBase64`/staged-deck parameter. Do **not** emit raw base64 PPTX in a normal command block;
use `slide` for ordinary one-slide creation unless the runtime has already staged a compiled deck
artifact.

**Advanced (`/`) when live-advertised:**

- `/insert-image` places a staged/approved image on a known slide or shape target. Inspect the slide
  first, supply explicit alt text, and avoid overlapping existing content unless requested.
- `/add-shape` creates a new shape on a known slide with explicit geometry; inspect existing layout
  bounds before choosing coordinates.
- `/format-shape` changes fill, line, text, size, or position on an inspected shape ref. Never infer
  the shape from visible text alone.
- `/add-table-slide` creates a slide with structured table rows. Use this for table-native slides
  instead of bullet text when the live grammar exposes it.
- `/apply-slide-layout` applies a discovered layout name/id to a known slide. List available layouts
  first and fail closed if the host does not expose the requested layout.
- `/insert-slide` with `deckBase64` is only for host-staged compiled deck artifacts; it is not a
  conversational shortcut for pasting arbitrary base64.
- Other modeled kinds (`/delete-slide`, `/move-slide`, `/duplicate-slide`, `/insert-hyperlink`) stay
  unavailable until the live grammar advertises them.

**Reversibility:** inserted slides and exact shape text replacements record bridge-level inverses.

**Gotchas:** PowerPoint's JS API is the narrowest — there is **no chart-from-data**, **no SmartArt**,
**no speaker-notes writer** (`set-speaker-notes` stays unadvertised), and **no loose-OOXML insert**.
Image/table/shape creation beyond those two writes stays modeled until the bridge has a tested host
write path and inverse.
