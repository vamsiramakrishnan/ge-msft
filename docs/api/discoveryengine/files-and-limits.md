# Files, context budget, and code execution

How much content to attach, by which mechanism, and when to hand a spreadsheet to **code
execution** instead of inlining it. Grounded in the `v1alpha`/`v1` discovery docs and the published
quotas; app-tier numbers are marked **indicative** (verify against the tenant's live quota).

## Two ways to provide content

| | **Inline context** (`query.parts[]`) | **Session context file** (`addContextFile`) |
|---|---|---|
| API | `v1alpha` streamAssist `query.parts[]` | `v1` `…/sessions/{session}:addContextFile` |
| Payload | `text` / document / drive / person refs | `{ fileName, mimeType, fileContents (base64) }` → `fileId` |
| Stored | not stored (per-turn) | stored on the **session** as `FileMetadata` (`fileId`, `mimeType`, `byteSize`, `tokenCount`, `downloadUri`); listable via `sessions.files.list` |
| Reasoning | grounding + citation over text/structure | grounding **and Python code execution** over the file |
| Best for | selections, ranges-as-tables, references — token-cheap, citable, structure-preserved | large or **analytical** files (xlsx/csv), multi-step compute, charts |
| Limit driver | model **context window** (token budget) | per-file size + session token budget; code-exec runtime |

`v1` streamAssist note: *"Empty query is only supported if `file_ids` are provided — the answer is
generated based on those context files."* So you can upload a spreadsheet and ask about it with no
other query text; the file *is* the input.

## Code execution (the xlsx path)

The assistant can run **Python** as part of answering — `AssistantContent.executableCode` (`code`,
"Currently only supports Python") and `AssistantContent.codeExecutionResult` (`outcome`, plus
`output` = stdout/stderr). It is enabled in the **assistant/engine config** (a tool), not via a
streamAssist request flag. With it on and a spreadsheet in session context, the assistant writes and
runs pandas-style code to compute aggregates, pivots, and charts.

- **Supported analytical formats:** XLS, XLSX, CSV, TSV (and Google Sheets in Workspace contexts).
- **Runtime constraint:** the Python sandbox has a **~30-second runtime limit** — a multi-join over
  ~100k rows can exceed it. For very large sheets, pre-aggregate in Excel, sample, or scope the range.

## Limits

**Verified (GE quotas page + discovery doc):**
- Data stores **100** / project (hard max **500**); engines **150** (max **500**); documents
  **10,000,000** per region.
- `completeQuery` **300/min**/project; search **300/min**/region.
- `checkGrounding.answerCandidate` ≤ **4096 tokens**; `semantic-ranker` input ≤ **512 tokens**;
  engine ingest chunking **100–500 tokens**/chunk.
- Each context file tracks `byteSize` + `tokenCount` → the session has a **token budget**, not just a
  byte budget.

**Indicative (Gemini app-tier / common docs — confirm per tenant, not API-guaranteed):**
- ~**10 files** per prompt; ~**100 MB** per file (~**2 GB** video).
- ~**1,000,000-token** context window on the higher tier (≈ 1,500 pages / 30k lines of code); lower
  tiers far smaller (~32k).

Sources: [GE quotas](https://docs.cloud.google.com/gemini/enterprise/docs/quotas) ·
[Agent Search quotas](https://docs.cloud.google.com/generative-ai-app-builder/quotas) ·
[addContextFile](https://docs.cloud.google.com/gemini/enterprise/docs/reference/rest/v1/projects.locations.collections.engines.sessions/addContextFile) ·
[StreamAssist files](https://docs.cloud.google.com/gemini/enterprise/docs/get-answers-from-streamassist).

## Decision policy for the add-in

Pick the cheapest mechanism that serves the intent (encoded in `@ge/content/budget.ts`):

1. **Already indexed in a connector?** → **reference** it (`documentReference`). ACL-preserving,
   citations resolve, ~no token cost. (The reference-over-inline rule.)
2. **Small/medium + grounding/Q&A/citation?** → **inline** via `@ge/content` (native GFM tables,
   contextualized chunks), staying within the inline token budget.
3. **Large, or analytical (compute/pivot/chart) xlsx/csv?** → **upload as a session context file**
   (`addContextFile`) and let **code execution** do the work. Mind the ~30s runtime + token budget.
4. **Budgeting:** sum `tokensEstimate` across attached context; as it approaches the inline budget,
   shift items to references or file-upload rather than truncating silently. Surface the budget in
   the context tray.

## Implication for our packages
- `@ge/content` already produces token estimates + native tables (inline path) and reference-or-inline
  selection. It now also exposes a **budget/strategy helper** (`recommendStrategy`) that returns
  `inline | reference | upload-for-code-execution` from `{ tokensEstimate, indexed, analytical }`.
- The **upload-for-code-execution** path needs a small **`v1` client** (`addContextFile` +
  `sessions.files`) alongside the `v1alpha` streamAssist client — a future addition to
  `@ge/gemini-client`, kept behind the same `ResolvedContext`/session abstraction.
