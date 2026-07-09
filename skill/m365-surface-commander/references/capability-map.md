---
title: Capability Map
kind: reference
skill: m365-surface-commander
topics: [surfaces, read-commands, write-commands, limits]
load_when: Checking which commands each Office surface may advertise.
---

# What each app can read and write

The commands available depend on which Office app the user has open. Each turn you are told
the available commands; this table is the full set they are drawn from. Only emit a command
listed for the current app.

## Reads

| App        | `outline` |                   `read`                   |      `search`       | `list`/`inspect`/`properties`/`open` | Surface refs exposed                       |
| ---------- | :-------: | :----------------------------------------: | :-----------------: | :----------------------------------: | ------------------------------------------ |
| Word       |    yes    |            yes (whole/section)             |         yes         |                 yes                  | selection, paragraphs, comments, anchors   |
| Excel      |    yes    | yes (`read <A1\|range>`, up to ~10k cells) |         yes         |                 yes                  | ranges, tables, named ranges, comments     |
| PowerPoint |    yes    |           yes (`read <slide:N>`)           |   yes (≤8 slides)   |                 yes                  | slides, shapes, text boxes, charts         |
| OneNote    |    yes    |              yes (whole page)              | yes (≤8 paragraphs) |                 yes                  | pages, paragraphs, tables, images          |
| Outlook    |     —     |              yes (whole item)              |   yes (≤8 lines)    |                 yes                  | current item, thread, attachments, draft   |
| Teams      |     —     |           yes (whole transcript)           |   yes (≤8 lines)    |                 yes                  | transcript segments, messages, deep links  |

Reads are bounded and scoped to the open item/window. They return empty on bad input and
always frame document content as data. `context` is runtime-served: it returns a strategy for
inline/reference/upload/code-execution grounding and never uploads, runs code, or writes content by
itself. `list`, `inspect`, `properties`, `comments`, `attachments`, `tables`, `slides`,
`neighbors`, and `open` are also runtime-served. They operate on typed host refs supplied by the
bridge; `open` only navigates/selects in the host and is never a write or a send action.

`ls` and `find` are also available across surfaces as DocFs read commands. `ls <path>` lists what
exists under `/doc` (the live document) or `/work` (saved workspace artifacts); `find <path> [glob]`
locates an entry under the same tree by name pattern. Both are read-only, bounded, and return empty
on bad input — they never mutate the host or infer content that isn't there.

## Writes

| Command   | Change             | Word | Excel | PPT | OneNote | Outlook | Teams |
| --------- | ------------------ | :--: | :---: | :-: | :-----: | :-----: | :---: |
| `set`     | write a cell       |      |  yes  |     |         |         |       |
| `grid`    | write a cell grid  |      |  yes  |     |         |         |       |
| `format`  | format cells       |      |  yes  |     |         |         |       |
| `table`   | create a table     |      |  yes  |     |         |         |       |
| `chart`   | insert a chart     |      |  yes  |     |         |         |       |
| `cf`      | conditional format |      |  yes  |     |         |         |       |
| `spill`   | write a table grid |      |  yes  |     |         |         |       |
| `suggest` | tracked change     | yes  |       |     |         |         |       |
| `comment` | add a comment      | yes  |  yes  |     |         |         |       |
| `reply`   | reply to a comment | yes  |  yes  |     |         |         |       |
| `slide`   | insert a slide     |      |       | yes |         |         |       |
| `shape`   | replace shape text |      |       | yes |         |         |       |
| `page`    | append a page      |      |       |     |   yes   |         |       |
| `mail`    | reply to mail      |      |       |     |         |   yes   |       |
| `compose` | draft new mail     |      |       |     |         |   yes   |       |
| `post`    | post to a channel  |      |       |     |         |         |  yes  |
| `/insert-text` | direct text insert | yes | | | | | |
| `/replace-selection` | direct selection replacement | yes | | | | | |
| `/insert-ooxml` | direct rich OOXML insert | yes | | | | | |
| `/fill-content-control` | fill known content control | yes | | | | | |

Every change is previewed and approved before it is applied and recorded for traceability. Some
writes are reversible by a bridge inverse or app review mechanism; direct inserts such as
`/insert-text` and `/insert-ooxml` are gated as irreversible until the bridge has a durable inserted
range handle. Outlook and Teams writes are always staged for review — never auto-sent. `compose`
never fills in recipients.

The Excel `table`/`chart`/`cf`/`spill` writes are reversible by a recorded inverse (delete the
created object, clear the rule, restore prior values). `chart` and `table` read a **range** —
to visualize a computed result, first `spill` the composed table into a grid, then point
`table`/`chart` at the resulting range (see the table → grid sink in command-grammar.md).
Use `grid` instead of repeated `set` commands when the whole rectangular cell payload is already
known, such as generated schedules, seed tables, and CSV-shaped results.

## Structured-operation equivalents by app

| App        | Best native shape for bulk/structured output                                                                 |
| ---------- | ------------------------------------------------------------------------------------------------------------- |
| Excel      | `grid` for literal rectangles; `spill` for computed tables; then `table`, `chart`, `cf` over the range.       |
| Word       | `suggest` for paragraph-level surgical changes; `/insert-table`, `/insert-content-control`, `/insert-html`, and `/insert-ooxml` for structured insertions. |
| PowerPoint | `slide` for simple bullets; `/add-table-slide`, `/add-shape`, `/set-shape-text`, `/format-shape`, `/insert-image`, and `/insert-html`-derived image/slides for designed layouts. |
| OneNote    | `page` for new synthesized pages; `/add-outline`, `/insert-table`, `/append-rich-text`, and `/insert-image` for structured page regions. |
| Outlook    | `mail`/`compose` and `/set-body`/`/prepend-body` for draft bodies; `/add-attachment` for generated files. Nothing auto-sends. |
| Teams      | `post` for reviewable text; `/post-card` for structured Adaptive Cards; Graph post/update/reply kinds only when estate writes are advertised. |

When the generated payload is large enough that a model should not hand-type it, first ask `context
analytical full-scope upload-preferred code-execution-preferred`. The host may return file/code
context; the model consumes only that structured result, then uses the native shape above.

## Out of scope

- Sending mail, creating calendar events or tasks, and checking out files are not available as
  commands here.
- Uploading a context file is not a model command. The model can use `context upload-preferred` to
  request a host/user attachment decision, then consume only the structured file id the host returns.
- If the user asks for something no listed command covers, emit `done`; do not invent a command.
