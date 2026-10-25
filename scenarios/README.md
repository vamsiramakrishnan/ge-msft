# Live feature scenarios

Each scenario runs one request the way a user would: a simulated Office document, the real pane
logic (planner → confirm → execute → approve), and the **real Gemini Enterprise engine** from
`packages/web-shell/.env`. Every approval is accepted automatically. Test environments only.

```bash
gcloud auth login                       # once; the runner signs in to Gemini with this account
bun run scenarios                       # all scenarios
GE_SCENARIO=excel-format bun run scenarios   # only ids containing this text (comma-separated)
```

Results: a pass/fail table in the terminal, and `scenarios/.results/<time>/` with `summary.md` plus
one JSON per scenario: every Gemini request and reply, the steps the pane showed, applied changes,
errors, and the document afterwards.

## Writing a scenario

Add an object to `excel.json`, `word.json`, `powerpoint.json` or `outlook.json`:

```json
{
  "id": "excel-format-professional",
  "surface": "excel",
  "prompt": "can you format the selected range, to make it look professional",
  "clarification": "NA",
  "seed": { "sheets": [{ "name": "Sheet1", "origin": "A1", "values": [["ID", "Name"], ["1", "Asha"]] }],
            "selection": "Sheet1!A1:B2" },
  "expect": { "applied": { "format-cells": 1 }, "formatsInclude": ["fontColor"] }
}
```

| Field | Meaning |
| --- | --- |
| `prompt` | Exactly what you would type in the pane (`/rewrite …` works too) |
| `clarification` | The answer if the planner asks a question |
| `approve` | `false` to reject every gate (checks nothing is written without approval) |
| `timeoutMs` | Per-scenario limit (default 240000) |

**Seeds per app**
- **Excel:** `sheets` (`name`, `origin`, `values` grid), `selection`, optional `comments` (`id`, `cell`, `content`, `replies`, `resolved`), `tables`, `namedRanges`
- **Word:** `paragraphs` (`text`, `styleBuiltIn` e.g. `"Normal"`, `"Heading1"`), `selectionText`, optional `comments` (`id`, `text`, `replies`, `resolved`)
- **PowerPoint:** `slides` (`id`, `shapes`: `text`, optional `type`, `placeholderType` e.g. `"Title"`), `selectedIndices`
- **Outlook:** `subject`, `body`, `from` (`displayName`, `emailAddress`)

**Expectations** (all optional; checked after the run)

| Key | Passes when |
| --- | --- |
| `status` | The last task ends with this status (default `completed`) |
| `noErrors` | No error is shown (default `true`) |
| `applied` | At least N changes of each kind were applied, e.g. `{ "write-cells": 1 }` |
| `notApplied` | None of these kinds were applied |
| `answerIncludes` | The final answer contains each text (case-insensitive) |
| `formatsInclude` | Excel: some formatted range has each property (`fontColor`, `horizontalAlignment`, `borders`, …) |
| `documentIncludes` | The document afterwards contains each text |
| `insertedSlidePositions` | PowerPoint: new slides occupy these exact one-based positions; all original stable slide IDs remain in their original order |
| `tableColumns` | PowerPoint: matching native table columns have only `allowed` cell values, contain required `includes` text, or have `nonEmpty` cells |
| `tableRowCount` | Tables matched by `tableColumns` have this exact number of data rows (header excluded) |
| `tableRows` | Tables matched by `tableColumns` contain each expected row once; string fields match substrings, array fields whitelist exact cell values |

Gemini's wording varies between runs, so check outcomes (what changed, final status), not exact text.
