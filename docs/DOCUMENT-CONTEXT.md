# Bounded Office document context

For a whole-source summary, ask **“read the entire doc and summarise it”**, **“summarize all 200
slides”**, or **“summarise the whole workbook”**. The runtime reads the open Office source to EOF,
summarizes bounded batches, reduces the partial summaries, then checks the source again before
publishing a final answer. A compact outline alone cannot satisfy this request.

## Office reads and model requests are separate budgets

| Surface | Office.js extraction | Scope and exclusions |
| --- | --- | --- |
| Word | `body.paragraphs.load({select: 'text,styleBuiltIn', skip, top: 41})`; consume 40 paragraphs and retain one lookahead | Main-body paragraphs, preserving empty paragraphs and order. Not rendered pages. Headers, footers, footnotes, comments, drawings and images are excluded. Table text is included only as exposed by body paragraphs; no table-structure guarantee. |
| PowerPoint | Slide IDs first; paged shape metadata, then up to 40 supported text frames in one `sync()` | Every slide, including empty slides; no 60-slide traversal ceiling. Speaker notes, grouped shapes, tables, charts, images and embedded objects are excluded and disclosed. Requires PowerPointApi 1.4. |
| Excel | All worksheet IDs/names/visibility; used-range dimensions; rectangular `getRangeByIndexes` reads of at most 1,024 cells, 32 columns and 64 rows | Values and formulas on every worksheet, including hidden sheets. Tiles advance across columns and then rows, respecting used-range offsets. Charts, images, comments and formatting are excluded. Requires ExcelApi 1.4. |

Word pages are layout-dependent: a 300-page document is traversed by paragraphs, not by an assumed
page count. Excel row count alone is also insufficient to predict cost: column count, text length
and formulas matter. One exceptionally large paragraph, shape or cell can still exceed a host
transport limit; the task reports failure rather than silently skipping it.

Ambient context is cheaper: Excel captures sheet/range metadata and at most one header row of
32 columns, after loading dimensions. PowerPoint captures the total slide count and at most
60 slides with eight shape previews each; shape metadata uses a lookahead and text reads use
batches of at most 40 objects. Word reads only a bounded paragraph prefix. Truncated slide,
worksheet, named-range and header inventories are labeled. Targeted questions and explicit attachments retain their
existing paths; this change does not turn every legacy search/attachment reader into a streaming
reader. In particular, Excel's legacy literal search/whole-sheet attachment can still load the active
used range, and PowerPoint's legacy search is limited to its existing slide prefix.

When ordinary chat finds no targeted evidence, its automatic body fallback uses one paged read
when available. It reports scope, extraction limitations, EOF and the next read command, with
explicit omissions when only part of that page fits. On Outlook and other nonpaged surfaces,
the fallback can resolve whole-item text, but automatic working context is capped at 24 KiB of
serialized, framed parts including its coverage notice. Oversized text keeps an exact Unicode-safe
prefix with an explicit partial marker. An explicit read and result inspection can retrieve omitted
text; that read observes the live source and is not a continuation of the earlier capture. These
ambient excerpts grant no source-completeness, freshness or write authority. Required evidence
bindings retain their separate fail-closed byte and freshness checks.

## Runtime and StreamAssist

1. `DocBridge.readDocumentPage({cursor, signal})` returns validated source units with stable-in-run
   IDs, a scope, extraction limitations, and an optional `nextCursor`. No cursor means EOF.
2. The runtime packs source fragments into at most 48 KiB of data per map request, including outer JSON escaping.
   Oversized units are split on Unicode code points, accounting for JSON escaping; text is not
   sliced away. The serialized query has a separate 96 KiB ceiling. The provider envelope adds
   fixed routing fields; these are byte budgets, not promises about provider token accounting.
3. Independent `isSessionLess: true` StreamAssist requests carry only the current source batch and
   task. The `document` route mounts no command/planner skill or client-selected search store.
   Existing conversation IDs, ambient selection text and old attached context are not replayed.
   Server-side engine configuration remains under tenant control.
4. Partial summaries are bounded to 12 KiB, with a requested length below 2,000 characters.
   Groups of up to four are reduced repeatedly. Source IDs and exceptions are requested at every
   stage. This is lossy semantic compression, not a proof that every fact survives the final summary.
5. The runtime re-reads every page and compares content hashes before publishing. A changed page,
   cursor cycle, duplicate source ID, budget overflow, failed host read, blocked model response or
   incomplete stream prevents a successful final summary. Cancellation stops subsequent calls.
   Office does not supply an atomic cross-host snapshot: two matching traversals detect drift,
   but cannot guarantee that no edit occurred after a page was checked.
6. The final answer includes deterministic coverage counts, scope and exclusions. Intermediate
   text and provider citation/provenance events are not published as final evidence. Source IDs
   requested from the model are textual references, not independently verified citations.

Default task limits: 256 model calls, 64 MiB of serialized source pages, 10,000 host pages and
1,000,000 distinct source-unit IDs. Hitting a limit reports an incomplete task; it never substitutes
an early-prefix summary. Repeated host traversal costs additional Office round trips. No full
source upload, local persisted index or background synchronization is introduced.

For exact spreadsheet totals, reconciliations or large structured comparisons, use the existing
Data workbench/`analyze` SQL path and send its bounded result to the model. A prose summary is not
a substitute for deterministic numerical computation.

## SDK and command usage

```ts
// Explicit scope avoids depending on the convenience English-language recognizer.
for await (const event of session.ask('Summarize the key decisions and exceptions', {
  scope: 'document',
  signal,
})) {
  // Render activity, the final token, done, or an explicit error.
}
```

`scope: 'attached'` uses the ordinary chat path, even when the wording mentions a whole document.
If a whole-source request also supplies `grounding.queryParts`, `dataStoreSpecs` or `fileIds`, it
fails with a scope-conflict error rather than silently ignoring those sources. The document route
reads the open file. Attached context already resident in the session is not part of that scope.

`AssistSessionOptions.documentSummary` exposes `maxCalls` and `maxSourceBytes`; both must be
positive safe integers. Increasing them increases total latency and provider usage, not request
size. The English recognizer is a convenience for common summary phrases, not a general intent
classifier. Unsupported surfaces return a clear error; they do not substitute an outline.

Bare `read` / `read document` now returns the first page of actual content. Follow the returned
`next` command (for example, `read document:40` in Word) until `eof: true`. `complete: true` means
that a single initial read reached EOF, not that the last page alone contains the whole source.
A pipeline cannot silently flatten an incomplete paged result into a complete document.
`outline` remains metadata. Natural-language requests beginning with “read” no longer execute as
arbitrary naked CLI selectors; explicit fenced CLI and addressed reads remain available.

Summary calls do not advance the existing server-side chat session. The answer is displayed in the
pane, but a later ordinary chat request cannot assume the provider remembers this isolated summary;
request the document again when a follow-up needs its contents.

## Validation and remaining release gates

Mocked Office tests traverse 200 slides (including a dense final slide), 6,000 Word paragraphs
(including blanks), and all cells in 30 worksheets of 1,500 rows with both 8- and 70-column shapes.
Those are extraction coverage tests, not measured live Office latency or a claim that every wide
workbook fits the default model-call budget. Runtime tests additionally cover 300 page-equivalents,
wire-byte limits, oversized Unicode text, EOF, cancellation, source drift, request isolation and
scope conflicts. Integration tests exercise the real StreamAssist request builder.

Live Word/Excel/PowerPoint requirement-set behavior, collection pagination and tenant StreamAssist
sessionless behavior still need acceptance in supported Office hosts. No live performance or
summary-quality result is claimed by the mocked tests.

## Reusable source workspace

The runtime now separates **source coverage**, **working context**, and **write authority**.
`EvidenceWorkspace` is shared infrastructure used by document summarization, source-bound session
execution, DocFs and the external retrieval pipeline. It is independent of a task verb, Office
surface, and model provider. Retrieval snippets are explicitly scoped to the snippets returned;
they are never advertised as complete external documents.

An Office object model supplies the structure analogous to a repository syntax tree. It is not a
lossless AST: extraction excludes unsupported objects as described above. Native Word anchors,
PowerPoint slide/shape identities and Excel addresses remain the bridge's responsibility. Snapshot
unit IDs are identifiers within a captured revision, not guaranteed durable live editing addresses.

| Repository operation | Source workspace operation | Result |
| --- | --- | --- |
| Directory/symbol browsing | `query(id, {kind: 'list', parentId?})` | Bounded node metadata and revision-bound continuation |
| Literal grep | `query(id, {kind: 'grep', literal, parentId?})` | Matching nodes and first UTF-16 match offset |
| Read exact source | `readText(id, unitId, {offset, maxBytes})` | Exact text window and `nextOffset`; no lost oversized nodes |
| Follow structure/references | `query(id, {kind: 'expand', id: unitId, direction})` | Parent, children, declared relations, or adjacent siblings |
| Build working context | `project(id, {requiredIds, query, ...})` | Exact required evidence, optional structure/retrieval, explicit omissions |
| Filesystem tools | `/evidence/<snapshot>/manifest.json` and `/evidence/<snapshot>/units/<group>/<ordinal>.json` | Existing DocFs read/list/search tools over immutable source data |
| Apply a surgical change | Existing native actuation with an evidence binding | Freshness and trusted target policy checked after approval, before dispatch |

DocFs unit files advertise `textWindowPath` before their text. Existing `cat`/read commands can
follow that path and each returned `nextPath` through every 4 KiB text window, including units larger
than the normal file-read budget. Window paths are virtual and need not be listed first.

`query` returns metadata, not full text. Its `limit` (maximum 100) and JSON byte budget (maximum
96 KiB) bound each reply. Continue until `complete`; cursors cannot cross queries or revisions.
`grep` is literal, not arbitrary regex execution. It returns one match per matching node; read the
node to inspect further occurrences. `readText` budgets text bytes; transport serialization and
other request fields require a separate envelope budget. `project` budgets the entire serialized
projection, including escaped source text, references and metadata. Required evidence that cannot
fit produces an error; it is never silently truncated. Lexical retrieval is a replaceable policy,
not a guarantee of semantic relevance or completeness.

```ts
const context = await session.prepareEvidence({
  requiredIds: selectedUnitIds,
  query: 'definitions assumptions exceptions',
  relationDepth: 2,
  neighbors: 1,
  maxBytes: 24 * 1024,
}, {
  // Trusted application code: validate native targets and operation parameters.
  // Omit this callback for read-only work. Evidence grants no write permission.
  authorizeEffect: request => selectedTargetPolicy.allows(request),
});
const snapshot = context.projection.snapshotId;
const hits = await session.sourceWorkspace.query(snapshot, {
  kind: 'grep', literal: 'exception',
});
// Follow hits.nextCursor, expand relevant nodes, and read exact text windows as needed.
for await (const event of session.runCommands('Revise the selected text consistently', {
  evidenceBinding: context.id,
  approvePlan,
})) render(event);
session.releaseEvidenceBinding(context.id);
await session.sourceWorkspace.release(snapshot); // explicit data retention lifecycle
```

`prepareEvidence` currently captures the full supported source when no snapshot is supplied.
Existing live `outline`, addressed reads/search and paged `read document:<cursor>` remain available
for lightweight exploration before deciding a full capture is needed. The snapshot navigation API
is **not** a new lazy live Office graph service. It does not force every ordinary task through a
whole-document scan. Applications can supply a narrower `EvidenceSource` with an honest scope, or
reuse a snapshot with `snapshotId` (which validates the live source before reuse).

The binding is process-local and never persisted with source data. Execution validates the source
at task start, again after approval immediately before a native effect, and before completing a
read-only task. A change outside the selected target can invalidate the result. The first native
write consumes the binding, including uncertain writes; a subsequent effect requires fresh context.
This deliberately does not infer a rebase or claim atomic compare-and-swap against concurrent Office
edits. Existing native anchor validation, approval, provenance, recovery and postcondition checks
remain in force. Streaming tokens can precede the final freshness check; only task completion is
withheld on stale evidence.

Bound model requests use isolated sessionless calls and append the explicit projection after
ordinary context compaction. The complete application request/context/grounding envelope is capped
at 96 KiB; a too-large required context fails explicitly. The provider never receives the entire
snapshot merely because it exists. Reading many pages over multiple turns is distinct from keeping
all pages simultaneously in the model's working context.

Source capture is atomic at validated EOF and content-addressed using SHA-256 (Web Crypto is
required; no degraded noncryptographic fallback). `MemoryEvidenceStore` defaults to eight snapshots
and 128 MiB, with explicit release rather than implicit eviction. The `EvidenceStore` port supports
persistent adapters; none ships in this change. Adapters must isolate principals, enforce retention
and aggregate quotas, and atomically publish immutable snapshots. Store reads are schema- and
hash-verified. The current adapter loads and verifies a whole snapshot for navigation; replacing
this with verified page/blob indexes is a future performance improvement, not a token-budget issue.

Explicit parent/related links support structural context. Current Office readers provide slide/shape
and tile/row parents; they do not infer Word section semantics, Excel's complete formula dependency
graph, or arbitrary cross-document links. Applications can enrich units and inject an
`EvidenceSelector` for semantic retrieval without changing capture, navigation or native execution.
The existing analysis workspace remains the place for bounded local SQL/computation on captured
tabular data. This change does not introduce arbitrary model-generated Office.js execution or claim
a new automatic snapshot-to-DuckDB importer.


The commander skill routes large-source and scoped-edit tasks to
`references/document-workspace.md`. Absolute-path `cat` and `grep` now dispatch through DocFs;
named artifacts and `ws:` handles retain WorkspaceStore behavior. SDK-only source capture,
structural queries and evidence-binding refresh are explicitly distinguished from CLI operations.
The planner keeps the edit scope fixed while describing broader evidence needs. Rebuild and validate
all skill bundles before a separately requested tenant provisioning step.
