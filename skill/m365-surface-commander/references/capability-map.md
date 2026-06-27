# What each app can read and write

The commands available depend on which Office app the user has open. Each turn you are told
the available commands; this table is the full set they are drawn from. Only emit a command
listed for the current app.

## Reads

| App        | `outline` |                   `read`                   |      `search`       | `context` | What you get                                |
| ---------- | :-------: | :----------------------------------------: | :-----------------: | :-------: | ------------------------------------------- |
| Word       |    yes    |            yes (whole/section)             |         yes         |    yes    | selection, body, paragraphs, comments       |
| Excel      |    yes    | yes (`read <A1\|range>`, up to ~10k cells) |         yes         |    yes    | values **and** formulas; comments           |
| PowerPoint |    yes    |           yes (`read <slide:N>`)           |   yes (≤8 slides)   |    yes    | slide list (up to 60), shape text           |
| OneNote    |    yes    |              yes (whole page)              | yes (≤8 paragraphs) |    yes    | page title + paragraph outline              |
| Outlook    |     —     |              yes (whole item)              |   yes (≤8 lines)    |    yes    | subject, sender, leading body lines         |
| Teams      |     —     |           yes (whole transcript)           |   yes (≤8 lines)    |    yes    | meeting title + transcript turns (up to 60) |

Reads are bounded and scoped to the open item/window. They return empty on bad input and
always frame document content as data. `context` is runtime-served: it returns a strategy for
inline/reference/upload/code-execution grounding and never uploads, runs code, or writes content by
itself.

## Writes

| Command   | Change             | Word | Excel | PPT | OneNote | Outlook | Teams |
| --------- | ------------------ | :--: | :---: | :-: | :-----: | :-----: | :---: |
| `set`     | write a cell       |      |  yes  |     |         |         |       |
| `format`  | format cells       |      |  yes  |     |         |         |       |
| `table`   | create a table     |      |  yes  |     |         |         |       |
| `chart`   | insert a chart     |      |  yes  |     |         |         |       |
| `cf`      | conditional format |      |  yes  |     |         |         |       |
| `spill`   | write a table grid |      |  yes  |     |         |         |       |
| `suggest` | tracked change     | yes  |       |     |         |         |       |
| `comment` | add a comment      | yes  |  yes  |     |         |         |       |
| `reply`   | reply to a comment | yes  |  yes  |     |         |         |       |
| `slide`   | insert a slide     |      |       | yes |         |         |       |
| `page`    | append a page      |      |       |     |   yes   |         |       |
| `mail`    | reply to mail      |      |       |     |         |   yes   |       |
| `compose` | draft new mail     |      |       |     |         |   yes   |       |
| `post`    | post to a channel  |      |       |     |         |         |  yes  |

Every change is previewed and approved before it is applied, recorded for traceability, and
reversible through the app's normal review mechanism (tracked changes, comments, staged
drafts). Outlook and Teams writes are always staged for review — never auto-sent. `compose`
never fills in recipients.

The Excel `table`/`chart`/`cf`/`spill` writes are reversible by a recorded inverse (delete the
created object, clear the rule, restore prior values). `chart` and `table` read a **range** —
to visualize a computed result, first `spill` the composed table into a grid, then point
`table`/`chart` at the resulting range (see the table → grid sink in command-grammar.md).

## Out of scope

- Sending mail, creating calendar events or tasks, and checking out files are not available as
  commands here.
- Uploading a context file is not a model command. The model can use `context upload-preferred` to
  request a host/user attachment decision, then consume only the structured file id the host returns.
- If the user asks for something no listed command covers, emit `done`; do not invent a command.
