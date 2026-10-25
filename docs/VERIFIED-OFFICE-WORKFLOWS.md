# Verified Office workflows: implementation and release boundary

This increment implements the shared execution foundation of the October improvement plan. It
does not declare the entire twelve-week roadmap complete. Native Office acceptance remains a
separate gate from model scenarios, simulator tests and browser-storage contracts.

## What is implemented

| Area | Implementation | Remaining boundary |
| --- | --- | --- |
| Host support | Current requirement sets, platform/build and read/compose mode filter Word, Excel, PowerPoint and Outlook discovery. Dispatch checks again. Effective manifests reach the task pane. | Member-specific optional variants still perform their own native checks. No preview API is promoted automatically. |
| Capability contracts | Four initial strict operation modules validate before approval and again at dispatch. A scaffold command reuses registry metadata. | Migration of the rest of the catalog is incremental. A descriptive module never installs a handler. |
| Evidence navigation | Authenticated page counts, page hashes, indexed unit reads and exact text windows. DocFs reads no longer load the full snapshot. Structural/literal queries scan at most eight stored pages per call. | Capture/complete-source verification still traverse the declared source. A full native dependency graph is not claimed. |
| Execution performance | Excel batches independent source captures into shared metadata/value stages. Analysis freshness checks use the port and correlate every returned source. | Other host plans and mutation batches retain their existing dependency ordering. No measured latency improvement is claimed before native benchmarks. |
| Write footprints | Common scalar formula shapes are checked before approval and in Excel. Unfamiliar or array output shapes are staged/refused; replacement of an existing unknown spill formula is refused. | This is conservative shape inference, not an Excel evaluator or a complete formula language. Bounded dynamic-array promotion needs its own native fixture. |
| Postconditions | Word tracked edits/content controls read current-revision text where supported. Word tracking mode is restored. PowerPoint shape writes read back text and retain inverses if verification fails. Excel formula verification waits for calculation Done. | Other native operations need operation-specific verifiers. A successful callback is not universal proof of the requested outcome. |
| Durable checkpoints | IndexedDB claims persist before dispatch, serialize across panes and block unresolved effects. Principal/document scopes are hashed. Saved approvals are never replayed. | This is a durable effect journal, not automatic replay of a whole saved DAG. Reconciliation of uncertain effects remains explicit. Unsaved files lack a durable document binding. |
| Managed output refresh | Word content controls and PowerPoint text shapes use captured document/object identity and a text hash. Three-way comparison preserves manual edits. Verified merge bases survive journal pruning. | Full chart/layout refresh and cross-application orchestration are not implemented here. |
| Useful workflows | Contract-review and Outlook-response entry points use the existing model/approval loop. Workbook close reuses exact-decimal reconciliation. SDK metric bindings preserve supplied decimal strings and lineage metadata. | Model judgment is still evaluated as model behavior; the wrapper alone does not prove a contract review is substantively correct. |
| Deck artifacts | Native editable charts, configurable theme fonts/colors, aggregate content budgets and opt-in geometry/accessibility checks extend the existing compiler. | Master/layout ingestion, font overflow and native visual fidelity require further work and real-host acceptance. |
| DevEx and acceptance | Office lab CLI, host-profile doctor, capability scaffold, native fixture runner, commit-specific report validator and task-pane certification controls. Browser CI checks actual IndexedDB cross-pane transactions. | No native-client or Gemini tenant execution is represented as passing without its corresponding evidence. |

## Use the same executor

`prepareOfficeWorkflow` returns a typed model task or a deterministic command program.
`AssistSession.runOfficeWorkflow` uses `runCommands`/`runCommandProgram`, including current
capabilities, approval, hooks, evidence bindings, recovery and receipt validation. Workbook close
uses the existing analysis compiler. There is no additional write executor or arbitrary Office.js
execution channel.

```ts
for await (const event of session.runOfficeWorkflow(
  { id: 'workbook-close', inputs: {
    leftRange: 'Expected!A1:C100', rightRange: 'Actual!A1:C100',
    leftKey: 0, rightKey: 0, leftAmount: 1, rightAmount: 1,
    currency: 'USD', tolerance: '0.01', destination: 'Report!A1',
  } },
  { approvePlan: reviewExactEffects },
)) render(event);
```

Managed refresh requires a native `captureTextTarget(ref)`, not a model-invented hash. The
prepared program carries `expectedDocumentId` and `expectedTextHash`; the bridge checks both
before changing the target. Missing adapter support is a refusal. The same value, including
quotes, tabs, paragraph breaks, Unicode and code fences, survives command quoting as data.

`bindBoardPackMetrics` binds explicit artifact rows/columns into named template placeholders.
It preserves exact-decimal strings, rejects missing/truncated inputs and returns source lineage.
The caller must obtain artifacts from its authorized workspace and check source freshness; a
well-formed artifact ID or hash is not proof of authority or currentness.

## Journal semantics and limits

The shell creates a journal for the signed-in principal and current saved document/item identity.
Pinned Outlook item changes rebind that scope. A document switch during a task or approval is a
conflict. A persistence failure before dispatch prevents the write. A failed checkpoint after
dispatch keeps an unresolved claim, preserves the actual returned receipt and blocks another
write pending inspection.

The journal keeps up to 1,024 entries and 8 MiB per scope. Binary payloads, cell grids,
provenance bodies and inverse bodies are omitted; a fingerprint identifies the complete intent.
Bounded generated text for Word/PPT refresh is retained as a merge base. Pruning removes known
older outcomes while preserving each latest verified merge base and every unresolved claim.
If those protected entries exceed capacity, the operation fails before dispatch. The existing
recovery store owns full cell inverses. This journal cannot replay omitted payloads.

The task pane offers an explicit inspection acknowledgement. Stop other running panes and
reconcile the document first. A recent dispatch claim cannot be acknowledged for five minutes.
Acknowledgement records an operator decision; it neither proves native state nor restores a
saved approval. Office provides no universal compare-and-swap, cross-application transaction
or unlimited exactly-once guarantee. Document renames/identity changes require a fresh binding.

## Office lab

```sh
bun install --frozen-lockfile
bun run typecheck
bun run office:lab list excel
bun run office:doctor host-profile.json
bun run office:lab scaffold word insert-html html,target
```

Copy the `host` object from task-pane diagnostics into `host-profile.json`. Doctor distinguishes
unsupported requirement sets, read-only documents and compose-only actions. It is an offline
explanation of observed support, not an execution grant. The scaffold refuses to overwrite a
file or duplicate an existing module. Implement required fields, native budgets, preview,
verifier and compensation, then register the trusted bridge handler and a native fixture.

Diagnostics contain host/build requirements, bounded native dispatch/batch timing, actual sync
counts, effect IDs and journal summaries. Native traces exclude request arguments, bodies,
queries, host error messages and credentials. Sync counts are actual observations; they are not
provider token usage, billable cost or a forecast of Office latency.

## Native certification

1. Build from a clean, exact commit. Build diagnostics now include the full SHA and a dirty marker.
2. Open the add-in in a disposable Word file, workbook, presentation or Outlook item. Append
   `office-lab=1` to the task-pane URL before sideloading/loading the pane.
3. Open **Tools → Native Office certification**. Confirm the disposable destination; Excel needs
   an explicit 2×2 range. Review the ordinary approval card. No model call is made by the fixture.
4. Inspect the outcome, perform the stated cleanup and copy the JSON report. Combine reports
   into an array. The SDK exports `nativeFixture` and `certifyNativeFixture` for automation.
5. Check the exact release head and required surfaces:

```sh
bun run test:native:check reports.json <40-character-commit> word,excel,powerpoint,outlook
```

The checker rejects missing, failed, unavailable, wrong-fixture, unknown-build or wrong-commit
reports. A simulator without native platform/build diagnostics cannot produce a passing native
report. The first fixture set proves tracked text readback, grid readback, inserted-slide text
and review-form opening respectively. It does not certify the entire action catalog, native undo
coverage, mail delivery or all Office clients. Choose and publish required client profiles before
promoting further capabilities.

```sh
bunx playwright install --with-deps chromium
bun run test:browser-contracts
```

These browser tests exercise real IndexedDB transactions across isolated panes and reopen;
they are separate from native Office acceptance and from real-model simulator scenarios.

## Next implementation gates

The remaining plan includes native pivots/filters/validation/worksheet lifecycle with precise
footprints and inverses, richer Word structure, complete native dependency indexing,
template/master-aware board packs and chart refresh, durable whole-program continuation,
cross-application artifact coordination, and supported Outlook compose/send-event checks.
They remain explicit backlog work. Promote each through the same contract, approval and
native-certification boundary; do not advertise a catalog-only API or substitute a static grid
for a requested native object.
