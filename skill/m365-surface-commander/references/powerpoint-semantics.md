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

**Reading:** `outline` gives metadata; `read document` pages supported shape text across the deck.
Follow returned `next` commands to EOF, retaining exclusions for unsupported content. Read other
slides for context while preserving the selected slides/shapes as the edit target. Captured unit
IDs are revision-scoped evidence, not native shape refs; see
[document evidence workspace](document-workspace.md). Write by slide index or an existing shape ref. Use `read slide:<n>` (it lists each shape's id and type and marks the title), `list shape`, `properties <ref>`, `open <ref>`, and `inspect <ref>` to target one
shape; do not parse a shape id out of slide text.

**Slide references (`slide=`):** a slide-targeted write (`/add-shape`, `/add-table-slide`) takes the
slide id from `outline` (e.g. `slide=256#0`), a 1-based slide number (`slide=3`), or `slide=last` for
the deck's last slide — including one a preceding `slide "Title" …` just added in the same plan (if
that `slide` step fails, the `slide=last` write is skipped rather than landing on another slide). Do
not invent ids such as `s2` or `new-slide`; an unknown reference fails with the list of valid
slides. The one exception is `/add-table-slide slide=new title="…"`, which **creates** a new slide
for the table.

**Shape references:** shape commands name the slide and the shape: `shape pp:shape:1:title "FY26
Plan"`, `/format-shape slide=1 shape=title fill=#0F6CBD fontColor=#FFFFFF`. The shape is its exact
shape id or `title` (the slide's title placeholder, else its first text shape). There are no shape
numbers: PowerPoint shape ids are small numbers themselves, so `2` always means id 2.
Positions (`left`/`top`/`width`/`height`, in points) must stay inside the slide; an off-slide
position fails with the slide's size.
To put a title and bullets on a new slide, use `slide "Title" "bullet" …` alone: it writes into the
layout's title/body placeholders, or adds text boxes when the layout has none.
To fill an **existing** slide whose title/body placeholders are empty (a blank "Title and Content"
slide, "Click to add title / text"), write into those placeholders — `shape pp:shape:<n>:title
"Contents"` and `shape pp:shape:<n>:<bodyId> "1. Background\n2. Roadmap"` (one item per
line; each line becomes its own paragraph). Never add a text box over empty placeholders: it sits
on top of them, and the bridge refuses it.

**Progressive disclosure:** specialized PowerPoint writes are live capability-gated. Use only slash
commands advertised by the grammar/help for the turn, after `list`/`properties`/`open`/`inspect`
confirms the slide, shape, layout, or staged asset. Fail closed: if the command, ref, layout, or
payload is absent, stop at a readable plan/comment instead of fabricating ids or silently degrading
to an unsupported Office.js path.

**Core verbs:**

- `slide "Title" "bullet" …` inserts a slide (bullets can be a table expression whose rows become
  bullets). It goes at the end unless you add `at=N`, which makes it slide N: "add a slide as the
  second slide" → `slide "Title" "bullet" at=2`; "after slide 3" → `at=4`. After an `at=` slide,
  address it as `slide=N`, not `slide=last`.
- `shape <pp:shape:slideId:shapeId> "text"` replaces text in one existing shape/text box. The deck
  outline lists each slide's shapes with their full ref (`shapes: pp:shape:256:2 Placeholder: SLA
  Terms; …`): use that ref as written, or `pp:shape:<slide>:title` for the slide title. For more detail use `list shape`, `properties <ref>`,
  `open <ref>`, and `inspect <ref>`.

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
- `/add-table-slide` adds a native table to a slide: `slide=new title="…"` creates a new slide for
  it (add `at=N` to make it slide N, as with `slide`); `slide=<id|number|last>` adds it to an
  existing one. Use it for table-native content instead of
  bullet text when the live grammar exposes it.
  For factual tables, preserve source precision and leave missing facts blank/TBD. Do not infer
  assignments or deadlines from the topic. Synthetic examples/proposals require an explicit user
  request and visible labelling; approval alone does not supply missing facts.
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
