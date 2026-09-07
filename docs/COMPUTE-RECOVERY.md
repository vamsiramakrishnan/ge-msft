# Data workbench, evidence and recovery

This implementation connects five reusable services to the existing task and approval lifecycle.
It ships a complete **Excel cell-analysis → reconciliation → preview → write → readback → undo**
path. Other Office surfaces retain their current bridges; versioned snapshots and durable undo
are not claimed for hosts that do not implement those ports.

## Use the workbench

Open **Data workbench** in Excel. Capture explicit source ranges, including their header row.
The resulting immutable tables retain worksheet identity, document identity, source address,
values/formulas hash, schema and lineage. They live in memory for the pane session.

Choose **Reconcile invoices and payments**, map key/amount/currency columns, and set a tolerance.
Alternatively supply one three-letter currency for both tables. The engine aggregates repeated keys
within each currency, uses DECIMAL(38,6), and performs a full outer join. It returns matched,
variance, unpaid, unallocated and invalid rows. Multiple payments are aggregated rather than
mislabelled as duplicate invoices. Currency groups are never silently combined. Six decimal places
are the precision policy; values with greater fractional precision are rounded by DuckDB's decimal
cast. Use a different explicit query if a different precision policy is required.

Finding chips use actual result counts. Clicking one rechecks source freshness and produces an
inspectable filtered table. A partial result remains partial through later queries and cannot be
written back. Source edits require capture and recomputation; hashes are freshness tokens, not
proof of authority or evidence of accounting correctness.

Enter an explicit **Write destination** and choose **Preview write**. The plan includes the exact
resolved grid and pinned destination. Approval, capability narrowing, lifecycle guards and host
policy gates still apply. The bridge checks source and destination versions immediately before
writing, then reads the result back. Literal result strings use Excel's text channel, preserving
large decimal strings and preventing formula injection. Formulas use a separate, screened channel.
A landed write remains `ok: true` if subsequent readback or checkpoint persistence fails; its separate
verification/recovery status prevents a false task-complete result.

## Model and extension APIs

The model, pasted command programs and the workbench use the same `AnalysisWorkspace`.
Model receipts contain bounded excerpts and handles, while the UI keeps its own structured preview. Run
`help analyze` for examples. The parser also has Python parity.

```text
analyze {"kind":"capture","range":"Invoices!A1:C100","headers":true}
analyze {"kind":"capture","range":"Payments!A1:C80","headers":true}
```

Use returned artifact IDs, never guessed IDs:

```text
analyze {"kind":"reconcile","spec":{"left":"INVOICES_ID","right":"PAYMENTS_ID","leftKey":0,"rightKey":0,"leftAmount":1,"rightAmount":1,"leftCurrency":2,"rightCurrency":2,"tolerance":"0.01"}}
analyze {"kind":"materialize","id":"RESULT_ID","destination":"Results!A1"}
```

Materialization enters the ordinary effect plan; it does not execute as a workspace read. Recovery
commands are restricted to explicit workbench user actions. SQL uses artifact table IDs and `c0`,
`c1`, etc. column names. `inspect`, `filter`, `remove` and `query` are also typed actions.

| Seam | Purpose | Implementation |
| --- | --- | --- |
| `DocBridge.captureCells` | Bounded, versioned host snapshot | Excel bridge |
| `ArtifactStore` | Content addressing, immutable copies, quotas, lineage | `@ge/compute` |
| `ComputeEngine` | Query/dispose port; lazy worker initialization | `@ge/compute/browser` |
| `AnalysisWorkspace` | Freshness, query/reconciliation, derived offers, materialization | `@ge/runtime` |
| `RecoveryCoordinator` | Checkpoint, reconcile, preview resume/undo, retention | `@ge/runtime` |
| `EvidencePipeline` | Selected-scope Search → Rank → framed excerpts → Grounding | Runtime hooks |
| `AssistSession.runAnalysis` | Task ownership, hooks, approval and execution | `@ge/runtime` |

Trusted extensions still register through `RuntimeExtension`. `message:received` now exposes
structured `dataStoreSpecs`; model/document content cannot register executable handlers. The
production evidence provider only adds a search when a request explicitly selects data stores.
Search requests preserve filters, returned resources must belong to those stores, and ranking only
reorders original excerpts. Excerpts remain untrusted context. Optional service failure is shown;
`requiredSupport` is a trusted tenant-level policy that blocks on missing/insufficient support.
Scores describe support in retrieved excerpts, not calculation accuracy or completeness. Other
existing attached/ambient context and normal streamAssist grounding are unchanged.

## Recovery and undo

Cell-write intent is saved **before** calling Office. Receipts and previous cells are stored in
Office document settings, scoped by the signed-in identity, and covered by the document's access
permissions. They contain no credentials or model conversation. This is deliberately separate from
the memory-only analysis store and metadata-only diagnostic ledger.

**Refresh recovery** compares an interrupted receipt with current cells. It identifies applied,
not-applied, conflict or uncertain state without replaying anything. A resumable write obtains a
new change ID, rechecks all sources, and requires a new approval. A supported undo restores previous
values and screened formulas only if the current destination matches the verified after-version.
It follows the same plan, hooks, gates and readback path. Unsupported/unsafe formula restoration
fails closed. Undo currently covers cell contents, not formatting, comments, charts, workbook
structure or every inverse descriptor advertised elsewhere.

The journal holds at most 32 effects and 2 MiB. Remove resolved receipts explicitly to free space;
unresolved receipts are not silently evicted. Same-origin panes serialize through Web Locks where
available, with an in-process queue otherwise. **Office settings are not a cross-device transaction
log or compare-and-swap service.** Coauthor edits can race the pre-write check; readback detects
mismatch but does not provide atomic isolation. No exactly-once execution claim is made. Unsaved
workbooks do not reuse a recovery identity across pane lifetimes. Renamed/unavailable worksheets
may need manual recovery.

Formula-looking literal cells are distinguished with native formula-area reads when needed
(ExcelApi 1.9). An older host that cannot make this distinction refuses the ambiguous snapshot.

## Compute/deployment limits

- Pinned DuckDB-WASM 1.32.0 and Arrow 17.0.0; lazy same-origin worker and WASM assets.
- Single SELECT or non-recursive WITH, analytical function allowlist, admitted artifact tables only.
- External access, automatic extension installation/loading and unsigned extensions are disabled;
  engine configuration is locked. No user JavaScript, Python, shell, URLs or arbitrary files.
- 100,000 cells / 256 columns per snapshot; 16 MiB per artifact; 32 artifacts / 32 MiB workspace.
- 30-second initialization, 10-second query budget; cancellation terminates the worker and resets it.
- DuckDB buffer-memory setting: 128 MiB. This is **not** a hard cap on total browser/WASM heap.
- Default 5,000 result rows, at most 100,000 result cells and 8 MiB. Truncation is explicit.
- Worker assets ship under `/compute/`. Vite and Firebase apply the worker-specific CSP:
  `default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; connect-src blob:; worker-src 'none'`.
  Other hosting systems must preserve this header. The main page fetches the same-origin WASM
  asset, then the worker loads only its blob URL. No CDN is required at runtime.

## Verification and bugbash

`bun run test` includes actual DuckDB-WASM execution of decimal reconciliation and external-access
controls; worker supervision tests; source/target freshness, failed checkpoints, interrupted
receipts, replay approvals and undo conflicts; scoped evidence/ranking/grounding failures; and
rendered workbench/controller state tests. The existing Office simulator integration suite remains
in place.

The bugbash also fixes late attach-after-detach, failed/policy-blocked chat command recovery,
throwing view observers interrupting settlement, PowerPoint evidence-review intent filtering,
partial-result lineage losing its truncation marker, cancelled tasks reaching Office after a slow
checkpoint, and dependency failures during pre-approval preparation.

For manual browser/Office validation, start the preview server and open `/analysis-preview.html`.
That separate development entry uses the real runtime and WASM with a labelled simulated Excel
host. It includes source-edit and checkpoint-failure controls and supports reload recovery. It is
not a production bundle entry and makes no model calls. Live Office behavior, tenant search/grounding
access, and deployed worker/CSP startup still require their actual host environments. In this
implementation session, cloud-browser access to the local test server was unavailable; rendered UI
and host-simulator tests are not represented as live Office or browser validation.
