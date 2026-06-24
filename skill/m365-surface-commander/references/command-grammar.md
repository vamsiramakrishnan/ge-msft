# Command grammar (full reference)

A small set of commands shared across all six Office apps. Only the **selector** (how you
point at content) changes per app. Read commands return content; write commands produce a
reviewable change.

## Selectors (the per-app part)

| App        | Selector                                         | Example                  |
| ---------- | ------------------------------------------------ | ------------------------ |
| Excel      | A1 range or named range                          | `Sales!C2:C7`, `Revenue` |
| Word       | exact text to anchor on (re-found at apply time) | `"Q3 revenue grew 12%"`  |
| PowerPoint | slide or shape                                   | `slide:4`                |
| OneNote    | page / whole                                     | (whole page)             |
| Outlook    | the open mail item (whole)                       | (whole item)             |
| Teams      | the transcript window (whole)                    | (whole transcript)       |

## Read commands

| Command   | Usage             | Notes                                                               |
| --------- | ----------------- | ------------------------------------------------------------------- |
| `outline` | `outline`         | Structure of the document/workbook. Not available in Outlook/Teams. |
| `read`    | `read <selector>` | Excel: an addressable range. Others: whole or current section.      |
| `search`  | `search <text>`   | Find content containing the text.                                   |

## Write commands

Each write command produces one reviewable change. A command is only available in apps
that support it (see [capability-map.md](capability-map.md)).

| Command   | Effect             | Apps        | Usage                                                               |
| --------- | ------------------ | ----------- | ------------------------------------------------------------------- |
| `set`     | write a cell       | Excel       | `set <A1> <value\|=formula>` — e.g. `set Sales!F2 =C2-D2`           |
| `suggest` | tracked change     | Word        | `suggest "old text" => "new text"` (anchored on exact text)         |
| `comment` | add a comment      | Word, Excel | `comment <cell> "text"` or `comment "anchor" "text"`                |
| `format`  | format cells       | Excel       | `format <range> k=v …` — keys: `bold italic fill numberFormat`      |
| `reply`   | reply to a comment | Word, Excel | `reply <commentId> "text"`                                          |
| `slide`   | insert a slide     | PowerPoint  | `slide "Title" "bullet" …` or `slide "Title" ($rows \| select a,b)` |
| `page`    | append a page      | OneNote     | `page "Title" "body"`                                               |
| `mail`    | stage a reply      | Outlook     | `mail "body"` — reviewable, never auto-sent                         |
| `compose` | draft a new email  | Outlook     | `compose "Subject" "body"` — recipients left to the user            |
| `post`    | stage a chat post  | Teams       | `post "text"` — reviewable, never auto-sent                         |

## Control commands

| Command | Meaning                      |
| ------- | ---------------------------- |
| `done`  | The whole task is complete.  |
| `help`  | List the available commands. |

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

## Reusable named commands

Define a parameterized command once, then call it by name. A name cannot reuse a built-in
command; an argument can only fill a declared `$param` (it cannot inject a new line).

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
