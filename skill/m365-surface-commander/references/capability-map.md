# What each app can read and write

The commands available depend on which Office app the user has open. Each turn you are told
the available commands; this table is the full set they are drawn from. Only emit a command
listed for the current app.

## Reads

| App        | `outline` |                   `read`                   |      `search`       | What you get                                |
| ---------- | :-------: | :----------------------------------------: | :-----------------: | ------------------------------------------- |
| Word       |    yes    |            yes (whole/section)             |         yes         | selection, body, paragraphs, comments       |
| Excel      |    yes    | yes (`read <A1\|range>`, up to ~10k cells) |         yes         | values **and** formulas; comments           |
| PowerPoint |    yes    |           yes (`read <slide:N>`)           |   yes (≤8 slides)   | slide list (up to 60), shape text           |
| OneNote    |    yes    |              yes (whole page)              | yes (≤8 paragraphs) | page title + paragraph outline              |
| Outlook    |     —     |              yes (whole item)              |   yes (≤8 lines)    | subject, sender, leading body lines         |
| Teams      |     —     |           yes (whole transcript)           |   yes (≤8 lines)    | meeting title + transcript turns (up to 60) |

Reads are bounded and scoped to the open item/window. They return empty on bad input and
always frame document content as data.

## Writes

| Command   | Change             | Word | Excel | PPT | OneNote | Outlook | Teams |
| --------- | ------------------ | :--: | :---: | :-: | :-----: | :-----: | :---: |
| `set`     | write a cell       |      |  yes  |     |         |         |       |
| `format`  | format cells       |      |  yes  |     |         |         |       |
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

## Out of scope

- Sending mail, creating calendar events or tasks, and uploading/checking out files are not
  available as commands here.
- If the user asks for something no listed command covers, explain it in plain text — do not
  invent a command.
