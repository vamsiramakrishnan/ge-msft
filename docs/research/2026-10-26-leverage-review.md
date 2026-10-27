# Leverage review — grammar, validation, composition, reliability, Office.js

Date: 2026-10-26. Scope: what would make the system materially better than it is today, ranked by
leverage. Every claim below was checked against the current head (`ec9aef9`) and, where it is an
external fact, against the Microsoft Learn requirement-set pages as of September 2026.

Shipping API levels today: WordApi 1.9 + WordApiDesktop 1.5, ExcelApi 1.21 + ExcelApiDesktop 1.1,
PowerPointApi 1.10, Mailbox 1.16. The bridges gate at most on WordApi 1.6, ExcelApi 1.12,
PowerPointApi 1.10 (selectively), Mailbox 1.15. The unified manifest requires Mailbox 1.12,
WordApi 1.3, ExcelApi 1.7 and declares no PowerPointApi requirement at all
(`manifests/m365-unified.manifest.json:71-73`).

## The ten highest-leverage moves

| # | Move | Why it is leverage | Effort |
|---|------|--------------------|--------|
| 1 | **One grammar authority.** Generate the Python preflight's argument grammar from the TS parser (or run the TS parser in the skill sandbox) and make the golden corpus a *differential* test TS≡Python. | Seven confirmed TS/Python divergences exist today; the "generated" Python table is a 29-line re-export, and the corpus is only run against Python. | M |
| 2 | **Structured parse errors + static typing.** Errors carry `{code, line, col, span, hint}`; a typing pass over `let` bindings and transforms runs before any host read. | Today errors are bare strings with no position, type errors surface only after the host read, and the repair budget is 2 turns. Better errors directly cut repair turns. | M |
| 3 | **Make the DAG a scheduler.** Run independent effects concurrently inside one `X.run`, honor `continue-independent`, and split approval by `approvalClass`. | The DAG exists but is only a skip-list; execution is a `for` loop; one `approvePlan` covers in-document, external and estate effects together. | M–L |
| 4 | **Excel undo groups and calc suspension.** Wrap every Excel effect plan in `Excel.run({ mergeUndoGroup: true })` (ExcelApi 1.20) and use `suspendApiCalculationUntilNextSync` + one explicit recalc per plan. | Gives the user a real single-step undo for every Excel kind (today undo exists for two kinds), and removes the forced full `Recalculate` plus up to 12 polls per write. | S |
| 5 | **Verify end-of-stream truth.** Emit `done` only on a terminal answer state or closing `]`; otherwise surface `incomplete_response`. | `stream-assist.ts:509-510` yields `done` on any clean EOF, so the runtime's guard (`assist-session.ts:4021`) never fires on the real client. A truncated stream can look complete. | S |
| 6 | **Word annotations instead of comments for findings.** Use `Paragraph.insertAnnotations` / `CritiqueAnnotation.accept()` (WordApi 1.7–1.8) for review findings; use `Range.highlight()` for preview. | On-canvas accept/reject with no document mutation until accepted; matches the "reviewed, attributable, recoverable" invariant better than comments. | M |
| 7 | **Richer pure algebra.** Add literals, `and`/`or` predicates, `join`, `group`/`agg`, `distinct`, `rename`, `pivot`, string and date functions, and `each` over a table. | The current transform set is 12 single-predicate operators; `let $x = 5` is a parse error; most multi-step analysis has to fall through to DuckDB `analyze`. | M |
| 8 | **Turn-level resilience.** Retry a dropped or malformed stream once per turn, honor `Retry-After`, wire the existing `CircuitBreaker`, and add a deadline plus retry for transient host errors (`GeneralException`, `ActivityLimitReached`). | A single mid-stream failure ends a 12-turn task; there is no retry on transient Office.js errors and no timeout on any `X.run`. | S–M |
| 9 | **PowerPoint bindings + slide images.** Anchor shapes by `Binding` id (1.8) instead of index, and read back `slide.getImageAsBase64()` (1.8) after a write. | Stable anchors across edits plus a visual receipt; today PowerPoint readback exists for two kinds only. | S–M |
| 10 | **Measure what ADR-0008 promised.** First-pass parse rate, capability-violation rate, repair turns, across base / commander / commander+CLI. | None of these is measured; the eval has 14 cases; the "95% first-pass" figure is a target, not a measurement. | S |

## 1. Command grammar and syntax validation

### What exists
- Hand-written line-by-line parser in `packages/contracts/src/command-grammar.ts` (2,431 lines) plus
  `expr-grammar.ts` and `skill-grammar.ts`. No tokenizer; quoting handled per verb in ad hoc ways
  (`:491`, `:1121-1127`, `:1540-1577`).
- Errors are `{ error: string }` with no position (`:378-380`). Did-you-mean exists for verbs and
  `/kind`, not for transforms, props or context kinds.
- Python preflight (`skill/m365-surface-commander/scripts/parse_commands.py`, 1,137 lines) is a
  hand port; `generated_language.py` re-exports vocabulary only. The manifest has no operator
  signatures or limits despite ADR-0008 §4.

### Confirmed drift (probed)
- `comment A1 "x | y"` → Python reports an unknown transform; TS parses it.
- `let $t = read A1:B3|sum x` → Python types it `Table`; it is a `Number`.
- `| SUM x` → Python accepts; TS runtime fails (case-sensitive lookup, `compose.ts:499`).
- `outline extra`, `read | count` → accepted on one side, rejected on the other.
- `def let(a):` is accepted; `shadowsBuiltin` omits `let`/`def`/`end` and is case-sensitive
  (`skill-grammar.ts:146-152`), and recipe calls route before expressions.

### Recommendations
1. **Grammar as data.** Define each verb's argument grammar once (a small PEG or a declarative
   spec in `@ge/contracts`) and emit both the TS parser tables and the Python tables from it, as
   ADR-0008 §4 describes. If that is too large a step, ship the TS parser itself into the skill
   bundle via a single-file build and have `surface_cli.py` shell out to it when Node is present.
2. **Differential corpus.** Run `skill/golden-corpus.jsonl` through both parsers in one Vitest
   test and compare full ASTs, not verb names. Add `fast-check` generators for command lines (the
   dependency is already present).
3. **Structured diagnostics.** `CommandParseError` → `{code, message, line, column, span?, hint?,
   didYouMean?}`. Echo the failing source line back inside the `result` block; today parse errors
   are pushed without it (`assist-session.ts:3089-3096`).
4. **Static type pass.** Before dry-run, infer `Table | Number | Text | Boolean | Selector |
   RangeRef` for every binding and reject `sum | head`, unbound `$ghost` inside `set = (...)`, and
   column-law violations (`columns(p) ⊆ C`). The runtime `Value` only knows three kinds today
   (`compose.ts:21-24`).
5. **Lexical hygiene.** Trailing `#` comments, line continuation `\`, case-insensitive transform
   names, and a single quoting routine shared by every verb.

## 2. Composition

### What exists
- Pipelines are pure-only; effects are terminal and take `$var` or `( pipeline )`.
- Transforms: `filter select sum avg min max count sort head tail sed derive`, each single-argument
  (`compose.ts:140-151`). No literals, no `and`/`or`, no joins or grouping.
- `def … end` recipes expand textually, session-only (`assist-session.ts:1132`).
- The effect DAG is inferred from range overlap (`planning.ts:278-302`) and used only to skip
  dependents; `failurePolicy` is always `stop-dependents` (`:299`); `propagateFailure` is never
  called; execution is sequential (`assist-session.ts:3748-3802`).
- Named ranges, whole columns and multi-area ranges fail `RANGE_RE`, and a failed parse means "no
  dependency" rather than "assume dependency" (`planning.ts:29`, `:92`).

### Recommendations
1. **Scheduler, not skip-list.** Topologically batch independent effects into one `Excel.run` /
   `Word.run` per batch, run batches in dependency order, and keep `stop-dependents` as the
   default while allowing `continue-independent` per plan. Treat an unparseable resource as
   overlapping everything.
2. **Phase boundaries from ADR-0008 §5, enforced.** Split a plan into phases at approval-class
   changes, fresh-observation needs and cross-artifact handoffs; present one approval card per
   phase. The classes are already computed (`assist-session.ts:2813-2815`) but a single
   `approvePlan` covers them.
3. **Algebra growth** (see move 7). Keep effects terminal; grow the pure side so the model does not
   need DuckDB for a two-table lookup. Add `each $row in $t: <effect template>` with a hard
   expansion budget so a top-N report can mint N comments in one approved, bounded plan.
4. **Durable recipes.** Persist `def` recipes in the document key's settings (Excel) or custom XML
   (Word) under provenance, so a recipe survives the task and shows up as a `/recipe` in `help`.
5. **Verify phase as a first-class node.** `finish when=verified` exists, but the examples still
   teach "re-read in a new turn" (`example-session-excel.md:50-57`). Make the compiler append a
   verify node per effect that has readback, so the model never spends a turn on it.

## 3. Reliability

### Confirmed weak points
- `done` on any clean EOF (`gemini-client/src/stream-assist.ts:509-510`); the JSON array parser
  does not require a closing `]` and skips non-matching frames.
- Mid-stream throw ends the whole task; backoff applies to the initial POST only
  (`stream-assist.ts:546-550`). `CircuitBreaker` (`retry.ts:123`) is never instantiated. No
  `Retry-After` handling.
- Idempotency: `changeId` is a fresh UUID per compile; dedupe is a content key in per-task
  WeakMaps (`assist-session.ts:3386-3442`). No journal when `documentKey` is undefined (unsaved
  file), and a bind failure at task start is swallowed (`:810-814`).
- Undo: only Excel `write-cells` and `sort-range` (`recovery.ts:125`). Inverse receipts are stored
  and never read (`provenance-store.ts:19-20`). Word insert-text/ooxml/table/find-replace return
  `not-reversible` although the manifest says reversible.
- Bridges: no retry on transient host errors (explicit, `bridge-dispatch.ts:46-56`), no timeout
  on any `X.run`, Outlook writes and Excel provenance `saveAsync` are unbounded, Excel
  `worksheets.getItem` throws instead of null-object, Word custom XML parts accumulate with no
  pruning (`host-port.ts:832`).
- Sync counts: Word tracked change ≈ 5–6 syncs plus two more `Word.run`s; Excel write-cells
  8–20 syncs including up to 12 calculation polls; PowerPoint format-shape syncs once per facet.

### Recommendations
1. Terminal-state check on the stream; a turn-level retry (once) that re-sends the same projected
   context; `Retry-After`; wire the breaker; an overall deadline per turn.
2. `mergeUndoGroup` + `suspendApiCalculationUntilNextSync` + `calculationMode` awareness for Excel;
   `untrack()` on paged reads; chunk large writes; a per-batch timeout and one bounded retry on
   `GeneralException`/`ActivityLimitReached`/`InvalidOperationInCellEditMode`.
3. Stable idempotency key = hash(documentKey, revision, effect content), stored in the journal, so
   a resume after reload can prove "already landed" instead of asking the model to skip work.
4. Use the inverse receipts: a real `undo <changeId>` verb for every kind that returns one, and
   stop advertising reversibility the bridge does not deliver.
5. Prune Word provenance parts to one rolling part; move Excel provenance from `settings` to a
   hidden named item or custom XML so it survives copy-sheet.

## 4. Office.js: advanced APIs worth adopting, by host

Each item names the requirement set and the capability it unlocks. Gate at runtime with
`isSetSupported`; the unified manifest can also require desktop-only sets.

### Word
- **Annotations and critiques (WordApi 1.7, popups 1.8).** `Paragraph.insertAnnotations`,
  `CritiqueAnnotation.accept/reject`, `Document.onAnnotationClicked/Hovered/Inserted/Removed`,
  `onAnnotationPopupAction`. This is the native "finding with accept" the Word mockup wants, with
  no tracked change until the user accepts. Requires an M365 subscription.
- **`Range.highlight()` / `removeHighlight()` (1.8)** for a non-mutating preview of a planned
  change before approval.
- **`Document.search` with options (1.7)** and `Paragraph.getText({includeTextMarkedAsDeleted})`
  to make anchor re-resolution deterministic under tracked changes; today `addComment` takes
  `results.items[0]` without the chooser (`host-port.ts:813`).
- **Content-control types (1.9) and events (1.5):** combo box, drop-down, checkbox, date picker;
  `onContentControlChanged` for the "surgical regeneration" scope lock.
- **WordApiDesktop 1.1–1.3:** `Document.compare()` for a redline against a prior version,
  `importStylesFromJson`, building blocks, `CustomXmlPart.selectNodes/appendChildNode` for
  structured provenance, `Range.hyperlinks`, `Document.detectLanguage()`.
- **Fields, bookmarks, footnotes, headers/footers** are already listed `now` in
  `CAPABILITY-CATALOG.md` but not built.

### Excel
- **`Excel.run({ mergeUndoGroup: true })` (1.20):** one Ctrl+Z per approved plan.
- **`Worksheet.evaluate(formula)` (ExcelApiDesktop 1.1):** evaluate a proposed formula before
  writing it; show the value on the approval card. **`Range.showPrecedents/showDependents`** and
  `Range.getDependents/getPrecedents` (1.15) feed the existing dependency-impact feature natively.
- **`suspendApiCalculationUntilNextSync`, `suspendScreenUpdatingUntilNextSync`,
  `calculationMode`:** replace the forced `Recalculate` plus polling loop.
- **Notes (1.18)** as a lighter, non-threaded annotation; **cell controls / checkboxes (1.18)**
  for review checklists; **`RangeTextRun`** for partial-cell formatting of citations.
- **Linked entity data types (1.19, load events 1.21):** this is the deferred build-plan item 2.2
  ("entity cells"), now feasible client-direct.
- **Custom functions:** `=GE.ASK` exists in `web-shell/src/functions/ge-ask.ts` but is not declared
  in the manifest; declaring it (with `CustomFunctionManager` visibility, 1.20) gives a formula-
  level entry point and a shared runtime.
- **Events:** `onCalculated`, worksheet `onAdded`, table events; `onChanged` args beyond `source`
  are currently discarded, which loses the changed address for freshness checks.
- Prefer `getItemOrNullObject` for worksheets and `getSpecialCells` / `RangeAreas` for scanning
  instead of loading whole used ranges (`excel-bridge.ts:313-316`, `:503-509`).

### PowerPoint
- **Bindings (1.8):** stable ids for shapes the agent wrote; the anchor for later `set-shape-text`
  and for undo.
- **`Slide.getImageAsBase64` / `Shape.getImageAsBase64` (1.8, 1.10):** a visual receipt and a cheap
  way to show the user the result on the approval card.
- **Tables (1.8 create, 1.9 format):** `addTable` with `values` + `mergedAreas` instead of text
  boxes; the whole-deck reader notes tables need 1.8.
- **Tags (1.3)** for per-slide provenance (declared unsupported today).
- **1.10:** `altTextDescription`/`isDecorative` for accessibility, `setHyperlink`, slide
  background, `themeColorScheme` for brand-consistent inserts, `exportAsBase64Presentation` for
  the "attach whole deck" flow.
- Speaker notes and image insert were removed on purpose; revisit speaker notes via
  `slide.notesSlide` when it reaches a numbered set.

### Outlook
- **Ribbon contexts.** The manifest declares only `default`; add `mailRead` / `mailCompose` so the
  pane is reachable where the Outlook bridge actually works.
- **Smart Alerts options (Mailbox 1.14+):** `errorMessageMarkdown`, `cancelLabel`,
  `sendModeOverride`; `OnNewMessageCompose` for draft bootstrap.
- **`notificationMessages`** for non-blocking receipts; **`customProperties` / `internetHeaders`**
  for provenance (declared unsupported today); **sensitivity labels (1.13)** as a send-gate input.
- Add `isSetSupported('Mailbox','1.5')` before `addHandlerAsync(ItemChanged)`.

### Teams
- Add the real `@microsoft/teams-js` dependency; the injected `TeamsJsLike` shape (`app.openLink`,
  `chat.openConversation({message, card})`) needs checking against v2 signatures, and there is no
  `app.initialize` / `getContext` / meeting detail API.

### Manifest and platform
- Declare a PowerPointApi requirement; raise ExcelApi/WordApi minimums or keep runtime gating but
  document the matrix.
- A **shared runtime** (`lifetime: long`) so custom functions, the pane and ribbon commands share
  state; keyboard shortcuts for "ask about selection".
- **Add-in as a Copilot agent action** (unified manifest; Word/Excel/PowerPoint, preview): the
  existing handler tables are already the "actions"; this is the largest distribution lever on the
  table, gated on the feature leaving preview.

## 5. Skill bundles and evaluation

- `SKILL.md` is lean (76 lines). The laws live in references but `surface_cli` does not enforce
  them; `--surface` is parsed and ignored; the effect budget default is 8 vs the runtime's 30;
  recipes are rejected as unknown verbs by the preflight.
- Hand-written catalogues still exist (`capability-map.md`, `command-grammar.md`, `algebra.md`)
  and disagree with the manifest (context kinds, value types, undocumented `derive`/`sed`).
- Examples teach the older style (`**thought**` prose outside the fence, `done` plus re-read),
  contradicting `SKILL.md:17` and `planning-normal-form.md:103-104`; one example uses an unbound
  `$top3`.
- Recommended: generate the three catalogues, parse-check every example and pattern with
  `analyze` (not `parse_block`), align budgets, implement the ADR-0008 evaluation matrix in
  `skill/eval` against a replayable fake host, and add eval cases for laws, `analyze`, recipes and
  `finish`.

## 6. Suggested sequencing

**Tranche A (1–2 weeks, mostly contracts + gemini-client):** stream terminal-state check;
structured errors with line/col; differential corpus test; `mergeUndoGroup` + calc suspension in
Excel; `isSetSupported` guards and timeouts in bridges; manifest ribbon contexts + PowerPointApi
requirement.

**Tranche B (2–4 weeks, runtime):** scheduler with batched `X.run`, per-phase approval, static
typing pass, inverse-receipt-backed undo, stable idempotency keys, turn-level retry.

**Tranche C (4+ weeks, bridges + skills):** Word annotations/critiques, PowerPoint bindings and
image receipts, Excel evaluate/precedents and linked entities, the richer algebra and `each`,
durable recipes, the eval matrix, Copilot agent actions when out of preview.

Each item touching identity, provenance or gates still routes through the `security-reviewer`
agent before it is checked off, per `CLAUDE.md`.
