# Flagship workflow acceptance

`runFlagshipTenantSuite` runs four named workflows through actual `AssistSession` execution and
the real bridges over disposable, labelled Office simulators. It produces metadata-only
`TenantWorkflowReport` evidence. It does not connect to a native Office client, certify an
advertised host profile, or exercise Entra/NAA federation.

| Workflow | Execution | Outcome checked |
| --- | --- | --- |
| Contract review | Real commander route in `runOfficeWorkflow` | Exact tracked payment-clause replacement, current-revision readback and preserved confidentiality paragraph |
| Workbook close | Deterministic recipe, real DuckDB WASM | Reviewed exact-decimal reconciliation materialization and native grid readback, including a decimal beyond IEEE-754 integer precision |
| Board-pack refresh | Deterministic artifact binding and managed refresh | Captured text-shape update with exact metric precision; manual risk-note shape remains unchanged |
| Outlook response | Real commander route in `runOfficeWorkflow` | Exactly one reviewable form with the requested source fact; `review-opened` is required, never treated as mail delivery |

Word and Outlook need a tenant model. Workbook and PowerPoint cases run without inference. The
suite also runs separate deterministic approval-refusal and post-approval source-change canaries
through the same runtime and native Excel bridge. These are simulated-host safety checks; their
flags must never be interpreted as native coauthoring certification.

## Run

```sh
GE_TENANT_WORKFLOWS=1 \
GE_TENANT_WORKFLOW_REPORT=/tmp/tenant-workflows.json \
bunx vitest run packages/web-shell/src/scenarios/tenant-workflows.live.test.ts --testTimeout=1900000
```

The live entry point binds the report to `git rev-parse HEAD` and records whether the checkout is
dirty. `GE_TENANT_EXPECT_COMMIT` can pin the expected commit; a mismatch fails before execution.
Dirty checkouts do not spend live model calls. Missing provider configuration still runs the two
deterministic workflows and safety canaries, then reports `BLOCKED_EXTERNAL` for absent model
evidence. The Vitest case is skipped after writing that report; process exit alone is not proof
of tenant acceptance. The release collector must inspect the report status and scopes.

Supply `GE_PROJECT_NUMBER` (or `GE_PROJECT`), `GE_LOCATION`, `GE_ENGINE`,
`GE_SURFACE_COMMANDER_SKILL` and one explicitly supplied short-lived access token or token-file
setting: `GE_WIF_ACCESS_TOKEN`, `GE_ACCESS_TOKEN`, `GE_WIF_ACCESS_TOKEN_FILE` or
`GE_ACCESS_TOKEN_FILE`. Existing `VITE_GCP_PROJECT`, `VITE_GCP_LOCATION`, `VITE_GE_ENGINE` and
`VITE_GE_SURFACE_COMMANDER_SKILL` configuration can supply endpoint/route settings. No login,
`gcloud` invocation, service-account fallback or implicit global-region choice occurs. This
harness uses the public transport; widget transport requires its separate existing acceptance
path. It does not upload or replace skills.

Defaults are one repetition and twelve model HTTP calls across the complete matrix. Set
`GE_TENANT_WORKFLOW_REPETITIONS` from one to five and `GE_TENANT_WORKFLOW_MAX_CALLS` from one to
24. A quota/auth/transport refusal stops later model cases while deterministic cases continue.
The model-call budget includes HTTP retries. Each active attempt has a three-minute cancellation
deadline and a bounded event stream; Node analytics has its own worker-enforced query deadline.
Host callbacks that do not support cancellation retain their actual receipts.

## Evidence boundaries

Reports explicitly distinguish `providerScope: live`, `replay` and `unavailable`;
`hostScope` is always `simulated`. `identityScope` distinguishes an explicitly supplied developer
token from synthetic replay and unavailable access. Requested skill mounting is observed on the
wire; actual configured agent invocation must appear in provider response metadata. Configuration
alone does not satisfy that check. Missing invocation metadata is unavailable evidence, not proof
that the task was semantically wrong.

Reports contain exact commit, fixed workflow IDs, repetitions, outcome/safety flags, call counts,
closed issue codes and elapsed-time statistics. They contain no prompts, documents, values,
recipients, bearer tokens, resource names, native locations or raw provider error messages. Wire
metadata observation is bounded to one MiB per response. Reports are imported evidence summaries,
not cryptographic attestations. Replay can verify harness behavior but cannot satisfy live-model
or native-client certification.

## Node analytics

`@ge/compute/node` exposes `createNodeCompute` for Node-side acceptance and SDK consumers. It uses
the same SQL policy, Arrow conversion, artifact integrity checks, result bounds and locked
external-I/O-disabled DuckDB configuration as the browser adapter. The Node subpath is separate
from the browser entry point. It admits only the current request's declared artifacts and drops
previously seeded tables outside that request.

Synchronous WASM executes inside a Node worker thread. The parent enforces `maxQueryMs` after
worker startup; abort, timeout, disposal, worker failure or an invalid/cross-query receipt
terminate the worker. A separate 30-second startup deadline covers engine loading. Timeout and
abort reject only after worker termination, and a subsequent query starts a fresh isolated
engine with no retained tables. Failure to confirm termination makes that adapter terminal.
The parent event loop remains available while SQL runs, including an expensive aggregate or
cross join. There is no automatic query replay. No model, Office or credential dependency is
introduced into compute, and Node worker imports remain outside browser bundles.
