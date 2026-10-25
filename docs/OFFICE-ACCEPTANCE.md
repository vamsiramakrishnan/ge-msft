# Office acceptance and saved-task operations

This phase operationalizes the workflows merged in PR #16. It adds native certification cases,
repeatable tenant scenarios, measured comparison reports, a stricter campaign gate and a pane for
safe saved-task inspection. Passing simulator, browser or replay tests never certifies a native
Office build or tenant model route.

## Saved tasks in the pane

Open **Tools → Saved tasks** in a saved document. **Refresh tasks** lists task metadata for the
current account/document binding. **Inspect task** reads the exact stored program and checks its
declared sources and recorded native outputs. **Review remaining steps** repeats those checks and
uses the ordinary approval card. Saving, listing or opening the pane never executes a saved task.

The **Save exact commands** form creates one stage. It requires an explicit, initially unchecked
declaration that the commands are the user's literal instructions and do not depend on uncaptured
evidence. It does not label model output as literal input. Supported inspected operations are
finite literal cell writes and exact Word content controls / PowerPoint text shapes. Formula,
expression, host-read, overlapping output and unsupported native-operation programs refuse the
automatic inspection route.

SDK programs can instead declare `evidence: {kind: 'cell-sources', sources: [...]}` with at most
16 captured native source versions. Declaring sources is a trusted integration contract, not a
model-derived proof of complete evidence coverage. Legacy tasks without a declaration remain
inspection-only. Current output validation requires the correlated receipt's `afterHash` and
native document/object identity; a past success or matching task ID is insufficient.

Native read callbacks have a two-second limit within a shared assessment deadline, and support
cancellation. Mutations are never abandoned behind a timeout or automatically repeated. Account
or document changes invalidate inspection. Actual landed write receipts remain visible after
cancellation. The account's persisted stage claims still block unknown outcomes.

**Hide completed task** removes a verified task from the active list. It preserves its claims and
ID so it cannot be recreated and replayed. It does not remove document content or reclaim the
32-task / 8 MiB scope quota. Use explicit inspected reconciliation for unresolved work; hiding a
task cannot grant replay authority.

Approval previews show native control/shape IDs and distinguish body start/end from insertion
relative to an anchored match. The advanced command includes freshness pins and round-trips
through the same compiler. No new write executor is installed.

## Eleven native cases

```sh
bun run office:lab fixtures
bun run office:lab fixtures excel
```

| Surface | Canonical cases |
| --- | --- |
| Word | Tracked edit; native table readback |
| Excel | Grid, table, grouped SUM pivot, literal filter and list validation readback |
| PowerPoint | Inserted slide text; captured native layout readback |
| Outlook | Review form opening; current compose subject/body readback |

Run only in a saved disposable file/draft. Builders use fixed bounded seeds, explicit native
destinations and ordinary approvals. Setup and operation have separate approvals so the operation
captures the actual landed seed. Cleanup is explicit; no fixture sends mail or pretends cell undo
also removes a native filter/validation rule.

```ts
const fixture = nativeOperationFixture({
  id: 'excel-pivot-readback', disposable: true,
  sourceRange: 'Source!A1:B3', destinationRange: 'Report!A1:C6',
});
const report = await certifyNativeFixture(session, bridge, fixture, commit, {
  approvePlan: reviewExactEffects,
});
```

`nativeFixture(surface, range?)` constructs the original smoke cases. `nativeOperationFixture`
constructs P0 cases, including a complete captured layout for PowerPoint. Factory results and
their effect/mode arrays are frozen. A modified/deserialized command cannot be certified under a
canonical fixture ID. Reports validate actual approvals, receipt kind/change-ID correlation,
effect sequence, supported APIs, stable native profile/document and observed readback. The
content-free report omits source bodies, native object IDs, URLs and credentials. Sync counts use
live trace subscription rather than a truncated diagnostics ring.

These cases are targeted contracts. They do not certify every Office operation, all Word table
topology, substantive contract-review judgment, mail delivery or actual native on-send event
delivery. Those remain separate native acceptance work.

## Tenant and deterministic flagship scenarios

See [tenant workflow acceptance](TENANT-WORKFLOW-ACCEPTANCE.md) for explicit provider configuration.

```sh
GE_TENANT_WORKFLOW_REPORT=scenarios/.results/tenant-workflows/report.json \
  bun run test:office-tenant
```

The suite runs contract review, workbook close, board-pack refresh and Outlook response. Word and
Outlook require the configured live commander. The other two use deterministic runtime/compiler
paths, including real DuckDB. Office hosts in this runner are always labelled `simulated`.
Requested skill mounting does not establish observed invocation. Reports distinguish live,
replay and unavailable provider access, enforce call/repetition budgets and reject dirty live
execution. The developer-token route is not evidence of browser Entra/WIF federation.

A blocked suite writes its JSON report and skips its Vitest case. Inspect the JSON status;
process exit zero is not acceptance. In this workspace, absent tenant configuration produced two
deterministic passes, two unavailable model cases and zero model calls. No native host or tenant
execution is claimed by that result.

`@ge/compute/node` supplies the same compute contract for Node harnesses. DuckDB runs inside a
terminable worker. Parent query deadlines, abort and disposal terminate the worker before another
query can begin; the next query starts fresh. Canonical SQL/external-I/O policy, artifact hashes,
byte/row limits and engine memory settings remain enforced. Node worker startup has its own
bounded deadline. No Node module enters the browser compute export.

## Measured workflow comparisons

`runOfficeBenchmarkComparison` runs baseline and candidate through separate disposable sessions
with separate ordinary approvals. Trusted, bounded read-only probes define the scoped source,
initial/output hashes and observed target/approval/completion/freshness invariants. Probe and
evaluator IDs are public code identifiers, not document names. Missing observations remain
unavailable; they are never zero-defect counters.

The runner records actual ledger/model counts and native sync traces. Native sync attribution
requires explicitly exclusive process activity; contamination invalidates it. Provider request
input comes from an observed body-byte meter, separate from command query bytes. Zero baselines
yield no reduction claim. A failed, cancelled, mismatched or unverified outcome cannot pass.
Reports retain origin labels and no prompts, source rows, target IDs, receipt bodies or tokens.

## Campaign collection and release gate

```sh
mkdir -p scenarios/.results/acceptance
bun run office:campaign:init > scenarios/.results/acceptance/bundle.json
bun run office:campaign scenarios/.results/acceptance/bundle.json
```

The empty bundle deliberately returns `BLOCKED_EXTERNAL` (exit 2). Populate its plan with the
actual advertised native surface/platform/build/mode profiles. The plan defaults to at least
20 comparison attempts per profile and workflow, 95% first-pass success and 50% reduction targets
for positive native sync / observed inference-input baselines. A zero-to-zero metric is N/A;
zero-to-positive is a regression. Targets are checked per profile, so one faster client cannot
mask another client's regression.

Collect independently exported, exact-head reports:

```sh
bun run office:lab campaign-collect plan.json local-checks.json native-reports.json \
  tenant-report.json benchmark-reports.json > bundle.json
bun run office:campaign bundle.json <40-character-commit>
```

`OfficeAcceptanceBundleSchema` owns the format. Local check records must come from actual
exact-head checks, including an explicit security review. The collector preserves execution
origins, rejects unknown report types, duplicate comparison IDs, contradictory receipt/check
claims, mismatched inputs/outputs and oversized data. CLI imports are bounded before reading
large files; schema/JSON errors do not echo document contents.

Every advertised profile needs its native cases, and every model-required workflow needs observed
live routing. Native comparisons must match the declared host/build/mode, use the correct
workflow, and contain complete native outcomes, probes, invariants and measurements. Replay and
simulated-host comparisons cannot satisfy native targets. Live-model/simulated-host reports
establish model evidence only. Full native model-workflow outcomes still require native runs.

Exit codes are 0 `PASSED`, 1 `FAIL`, 2 `BLOCKED_EXTERNAL`. Imported metadata remains `evidenceOnly`;
it is not cryptographic attestation, a production promotion grant or automatic deployment. Native
Office and tenant execution, advertised-client certification and release benchmark targets remain
open until corresponding evidence is collected from the exact clean release commit.

## Implementation validation

The aggregate local suite passes 3,995 tests, with 26 optional default skips (22 earlier live/model
cases, the tenant matrix and three separately enabled browser tests). All three actual Chromium
storage tests pass, including hide/reopen and retained-ID replay refusal. TypeScript, lint/design,
production builds, generated artifacts, source-exact skill bundles and secret scan pass. Python
passes 76 shared parity cases, six manifest tests, 40 tooling tests, parser/checker self-tests and
14 deterministic evals. Local Bun is 1.3.14; CI remains pinned to 1.3.13.

Inspected 48 screenshots at 320/360/480 across ready/blocked/complete/empty/error/loading saved
tasks, command inspection/creation, and collapsed/expanded body, anchor, control and shape
approvals. Controls wrap without new horizontal overflow; advanced commands intentionally scroll
inside their code block. The pre-existing Settings rail label remains clipped. Browser previews
use explicitly labelled sample data and do not establish native Office acceptance.

Independent review has no unresolved high or medium findings after native-fixture immutability,
report coherence, read deadlines, worker termination and target-faithful approval fixes.
