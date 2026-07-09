---
title: Command Grammar
kind: reference
skill: m365-surface-commander
topics: [cmd-block, read-verbs, write-verbs, composition, recipes]
load_when: Exact CLI syntax, selector forms, composed writes, or reusable recipe grammar is needed.
---

# Command grammar (full reference)

A small set of commands shared across all six Office apps. Only the **selector** (how you
point at content) changes per app. Read commands return content; write commands produce a
reviewable change.

## Selectors (the per-app part)

| App        | Selector                                         | Example                  |
| ---------- | ------------------------------------------------ | ------------------------ |
| Excel      | A1 range, table, or named range                  | `Sales!C2:C7`, `Revenue` |
| Word       | exact text, paragraph, comment, or content control anchor | `"Q3 revenue grew 12%"`  |
| PowerPoint | slide, shape, text box, table, or chart          | `slide:4`, `shape:logo`  |
| OneNote    | page, paragraph, table, or image anchor          | `page:current`           |
| Outlook    | current item, thread, compose body, draft target, or attachment | `item:current`           |
| Teams      | message, thread, channel, transcript segment, or deep link | `thread:current`         |

## Read commands

| Command       | Usage                         | Notes                                                               |
| ------------- | ----------------------------- | ------------------------------------------------------------------- |
| `outline`     | `outline`                     | Structure of the document/workbook. Not available in Outlook/Teams. |
| `read`        | `read <selector>`             | Excel: an addressable range. Others: whole or current section.      |
| `search`      | `search <text>`               | Find content containing the text.                                   |
| `list`        | `list [kind]`                 | List addressable context refs. `kind` is optional.                  |
| `inspect`     | `inspect <refId\|selector>`   | Resolve one context ref or selector into readable content.          |
| `properties`  | `properties <refId\|selector>`| Return safe metadata: id, kind, title, locator, host ref.           |
| `comments`    | `comments [selector]`         | List comment refs, optionally scoped near a selector.               |
| `attachments` | `attachments [selector]`      | List attachment refs, optionally scoped near a selector.            |
| `tables`      | `tables [selector]`           | List table/range refs, optionally scoped near a selector.           |
| `slides`      | `slides [selector]`           | List slide refs, optionally scoped near a selector.                 |
| `neighbors`   | `neighbors [refId\|selector]` | Show nearby context refs around a target.                           |
| `context`     | `context [hints]`             | Ask the host for context/upload/code-execution strategy. Read-only. |
| `open`        | `open <refId\|selector>`      | Navigate/select in the host only. Never sends or mutates.           |

Context kinds accepted by `list`: `selection`, `range`, `comment`, `slide`, `message`,
`attachment`, `table`, `paragraph`, `shape`, `page`, `transcript`, `file`, `reference`.

Use `list`/`inspect`/`properties` before a surgical action when the snapshot only says
"selected range" or "current item". Use `open` only to help the user see the target in the host;
it is not an approval and cannot apply a change.

`context` accepts zero or more hints:

`incremental`, `inline-preferred`, `reference-preferred`, `upload-preferred`,
`code-execution-preferred`, `analytical`, `full-scope`.

Use `context` before escalating from bounded reads to full-file attachment or hosted analysis. It never
uploads by itself, never runs code, never approves a write, and never grants a capability.

```
context analytical full-scope upload-preferred code-execution-preferred
```

If the result recommends upload, wait for the host/user to attach the file and provide a structured
file id. Do not invent file ids or emit upload/code commands.

## Local workspace commands

Workspace commands create and inspect bounded virtual artifacts inside the add-in runtime. They are
the local workbench for large reads, derived notes, chart data, markdown tables, or cross-step
handoff packets. They are **not host writes**: they do not mutate Office content, upload files, run
code, send messages, or approve later effects.

| Command     | Usage                                                                 | Notes |
| ----------- | --------------------------------------------------------------------- | ----- |
| `workspace` | `workspace [name\|ws:id]`                                             | List artifacts, or summarize one artifact by name/id. |
| `save`      | `save <name> = read <selector>`                                       | Save a bounded rendered read/search/outline/pipeline/literal output. |
| `cat`       | `cat <name\|ws:id> [head=N]`                                          | Preview the first N lines of an artifact. |
| `grep`      | `grep <name\|ws:id> "pattern" [context=N]`                            | Search an artifact locally and return compact line matches. |

Accepted `save` sources:

```
save schedule.tsv = read 'Daily schedule'!B3:I53
save matches.md = search "termination"
save outline.md = outline
save note.md = "Data read from the current selection."
save top-regions.tsv = (read Sales!A1:D5000 | sort Revenue desc | head 10)
```

Use workspace artifacts when the same read will be inspected repeatedly or when a raw result would
make the pane noisy:

```
save schedule.tsv = read 'Daily schedule'!B3:I53
grep schedule.tsv "Manager Sync" context=1
cat schedule.tsv head=20
```

Staleness rule: an artifact is a snapshot of rendered content. If you mutate the host and need fresh
truth, `read`/`save` again. Current pipelines cannot start from a workspace artifact directly; use
`cat`/`grep` for inspection, or recompute a table from a live `read` when a later `spill`, `grid`,
`table`, or `chart` needs a concrete range/value.

## Write commands

Each write command produces one reviewable change. A command is only available in apps
that support it (see [capability-map.md](capability-map.md)).

| Command   | Effect             | Apps        | Usage                                                                                                 |
| --------- | ------------------ | ----------- | ----------------------------------------------------------------------------------------------------- |
| `set`     | write a cell       | Excel       | `set <A1> <value\|=formula>` — e.g. `set Sales!F2 =C2-D2`                                             |
| `grid`    | write a cell grid  | Excel       | `grid <range> = "a\tb\nc\td"` — write a rectangular literal TSV grid as one change                    |
| `suggest` | tracked change     | Word        | `suggest "old text" => "new text"` (anchored on exact text)                                           |
| `comment` | add a comment      | Word, Excel | `comment <cell> "text"` or `comment "anchor" "text"`                                                  |
| `format`  | format cells       | Excel       | `format <range> k=v …` — keys: `bold italic fill numberFormat`                                        |
| `reply`   | reply to a comment | Word, Excel | `reply <commentId> "text"`                                                                            |
| `slide`   | insert a slide     | PowerPoint  | `slide "Title" "bullet" …` or `slide "Title" ($rows \| select a,b)`                                   |
| `page`    | append a page      | OneNote     | `page "Title" "body"`                                                                                 |
| `mail`    | stage a reply      | Outlook     | `mail "body"` — reviewable, never auto-sent                                                           |
| `compose` | draft a new email  | Outlook     | `compose "Subject" "body"` — recipients left to the user                                              |
| `post`    | stage a chat post  | Teams       | `post "text"` — reviewable, never auto-sent                                                           |
| `table`   | create a table     | Excel       | `table <range> [headers] [name=NAME]` — promote a range to a native Table                             |
| `chart`   | insert a chart     | Excel       | `chart <type> <range> [title="…"] [series=rows\|columns]` — types: `column bar line pie scatter area` |
| `cf`      | conditional format | Excel       | `cf <range> >VALUE [fill=#hex]` · `cf <range> databar\|colorscale` · `cf <range> top=N`               |
| `shape`   | replace shape text | PowerPoint  | `shape <pp:shape:slideId:shapeId> "new text"`                                                        |
| `spill`   | write a table grid | Excel       | `spill <range> = (<table expr>)` — write a composed table as a cell grid                              |

PowerPoint generated-deck import is a client-staged artifact path: the host compiles a bounded
DeckSpec into one base64 `.pptx` and invokes `insert-slide` with `params.deck` after preview and
approval. Do not ask the model to write raw PPTX base64 in the CLI.

## Control commands

| Command | Meaning                      |
| ------- | ---------------------------- |
| `done`  | The whole task is complete.  |
| `help`  | List available commands, or `help <command>` for one generated playbook. |

`<command> -h` and `<command> --help` are aliases for `help <command>`. Use targeted help before
object-rich operations (`shape`, `chart`, `format`, specialized `/<kind>` commands) instead of
guessing a long syntax from memory. Targeted help is read-only/control and comes from the same
generated `m365-cli-1.0.json` manifest as the runtime grammar.

## Composition

Reads produce **values** (a table, a number, or text). Pure transforms combine values via
pipes (`|`) and named bindings (`let`). **Pipelines only read and compute — they never
write.**

```
read <selector> | filter <col><op><val> | sum <col>     -> a value
let $x = read <selector> | filter region=East           -> bind it; reuse as $x
$x | count                                              -> a $var can start a pipeline
```

Transforms:

| Transform                     | Usage                                                                        |
| ----------------------------- | ---------------------------------------------------------------------------- |
| `filter`                      | `filter <col><op><val>` — ops: `= != > < >= <= contains`; keep matching rows |
| `select`                      | `select <col,col,…>` — keep only these columns                               |
| `sum` / `avg` / `min` / `max` | `<agg> <col>` — over a numeric column → a number                             |
| `count`                       | `count` — number of rows → a number                                          |
| `sort`                        | `sort <col> [desc]` — sort rows by a column                                  |
| `head` / `tail`               | `head <n>` / `tail <n>` — first / last n rows                                |

### Composed writes

A write's value/text slot can consume a composed value — write `$var` or a parenthesized
pipeline `( … )`, and it is resolved to a literal before the change is applied. Everything
else is literal text. A pipeline cannot write (`$var | set …` is rejected).

```
let $a = read Sales!A1:B9 | filter region=East
set Summary!B2 = ($a | sum amount)
set B3 = $total
```

All writes in a turn are previewed together as one set of changes, approved once, then
applied and recorded one by one.

### Literal grids (Excel): generated tables and schedules

Use **`grid`** when you already have a complete rectangular literal table, such as generated
schedule seed data or CSV-shaped output. It compiles to one `write-cells` effect, so the user sees
one preview and one approval for the rectangle instead of repeated single-cell approvals.

```
grid 'Daily schedule'!C5:E7 = "Monday\tTuesday\tWednesday\nIndia Sync\tMusic Lesson\tIndia Sync\nDeep Work\tDeep Work\tDeep Work"
```

Rules:

- The body is a quoted TSV literal: use escaped `\t` for columns and escaped `\n` for rows.
- Rows must be rectangular. Ragged rows are rejected.
- The target range must name the rectangle being written. Prefer explicit ranges for legibility.
- Use `set` for one scalar cell.
- Use `spill` when the grid comes from a composed table expression, not literal generated data.
- For very large, algorithmic, or file-derived grids, ask `context analytical full-scope
  upload-preferred code-execution-preferred` first and wait for structured host context.

### The table → grid sink (Excel): analyze → shape → materialize → visualize

A pipeline value is usually a number or text. The **`spill`** verb is the one place a whole
**table** value lands in the document: it writes the table's rows as a **cell grid** (the dual
of `set`, which writes one scalar). That single sink turns "analyze → shape → materialize →
visualize" into one pure, gated pipeline — `spill` materializes the grid, then `table` and
`chart` consume the resulting **range** (never a table value directly, so the anchor stays a
concrete range).

```
# 1 · analyze + shape — bind the top 10 regions by revenue (pure; nothing written yet)
let $top = read Sales!A1:D5000 | filter Quarter=Q3 | select Region,Revenue | sort Revenue desc | head 10

# 2 · materialize — spill the composed table into a grid (the table → grid sink)
spill Report!A1 = ($top)

# 3 · promote the spilled grid to a native Table (headers in row 1)
table Report!A1:B11 headers

# 4 · visualize the same range as a chart
chart column Report!A1:B11 title="Top regions by revenue"
done
```

`spill`'s argument MUST be a table expression — a `$var` or a parenthesized pipeline
`( read … | select … )`. A literal is rejected: use `set` to write one cell, `spill` to write a
table. You can also select columns inline, e.g. `spill Report!A1 = ($top | select Region,Revenue)`.

## Recipes (reusable named commands)

A **recipe** is a parameterized command you define once with `def … end`, then call by name. (We call
these in-language definitions _recipes_ — "skill" is reserved for the Agent Skill bundle this file
lives in.) A recipe name cannot reuse a built-in command; an argument can only fill a declared
`$param` (it cannot inject a new line).

```
def reconcile($a $b):
  let $x = read $a | sum amount
  set $b = $x
end

reconcile Sales!A1:B9 Summary!B2
```

## The fenced block

Emit exactly one fenced ` ```cmd ` block per turn. Only the contents of that block
run; any reasoning text outside it is ignored. A turn with no fenced block is treated as a
prompt to try again, not an error.
