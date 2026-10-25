# Office capability phases: implementation and acceptance

This continues merged PR #15 on the same client-direct architecture. The supported code paths for
the remaining phases are implemented below. Native Office, tenant-model and measured release
acceptance remain open; simulator and browser results do not satisfy those gates.

| Phase | Delivered code | Acceptance boundary |
| --- | --- | --- |
| Capability contracts | Strict parameter modules cover every installed Word/Excel/PowerPoint/Outlook handler, including nested fields. Handler coverage is tested. Current host capabilities still control discovery and dispatch. | A module describes required postconditions; only an actual native verifier can establish them. |
| Bounded execution and evidence | Excel native formulas, names and tables feed a revision-bound dependency index; reverse impact queries reuse unchanged pages. Independent range capture shares sync stages. | Dynamic, external, whole-row/column and 3D references require full refresh. The two-range simulator fixture uses two syncs instead of four with identical snapshots; no native latency or general 50% claim is made. |
| Workbook close | Native grouped SUM pivots, literal-value filters, list validation, table promotion and worksheet create/rename/activate/protect use narrow footprints, preview-time pins and native readback. Existing exact-decimal reconciliation stays available. | Complex pivot dimensions, existing filters/validation, header-row insertion, worksheet deletion/unprotect and formula validation refuse. Pivot numeric verification uses Excel/JS finite numbers and relative tolerance, not financial exact-decimal arithmetic. |
| Managed board packs | Bounded native master/layout/placeholder capture, hash-pinned layout application, reusable compiler masters, artifact-bound text/charts, exact metric lineage and three-way artifact refresh. Existing reviewed PPTX import materializes editable charts. | PowerPoint's supported API here has no native chart-series editor. Refresh regenerates editable artifacts for reviewed import, without overwriting an existing chart. Template metadata is not a clone of every theme/image/master asset. Layout changes have no automatic undo. |
| Durable continuation | Persisted dependency stages, atomic claims, correlated effects, reopen validation and explicit inspected reconciliation. Browser transactions protect concurrent panes. | At most 32 stages, one host effect per stage. Stage-local bindings are recomputed. Crashes and unknown outcomes block replay; there is no expiring retry lease or exactly-once Office transaction. |
| Cross-application work | A typed coordinator binds fresh Excel metrics to Word controls, PowerPoint text and a final Outlook review form through registered authenticated session ports. Every destination uses its existing approvals, journal and executor. | Hosts must be open and registered by trusted integration code. This SDK does not launch/control another Office process or provide a remote host transport. Each destination commits independently; there is no cross-app rollback transaction. |
| Word and Outlook | Native bounded table windows, inserted-table readback, bounded draft-field verification, complete compose facts and deterministic send policy in the actual function runtime. | Table windows omit nested/merge topology. A review form is not sent mail. Send rules are explicit deterministic checks, not substantive grounding/DLP certification. |
| Release and DevEx | Exact command help, generated skill parity, authenticated SDK ports, native lab and deterministic readiness evaluation. | Release requires exact-head native client profiles, tenant workflow outcomes and performance evidence. Missing external evidence is `BLOCKED_EXTERNAL`, never a pass. |

## Native Excel commands

The runtime captures native source/destination hashes and worksheet identities during dry-run,
before approval. Supplied pins are checked rather than overwritten. Native handlers check them
again before mutation. Direct adapter calls must supply their own captured pins. No link or hash
grants authority over another range.

```text
/insert-pivot sourceRange=Sales!A1:C20 destinationRange=Summary!A1:D30 rowFields=Region valueFields=Revenue
/filter-range range=Sales!A1:C20 column=Region criterion1=APAC
/set-data-validation range=Sales!D2:D20 values="Open,Closed"
/manage-worksheet action=rename name=Summary newName=Close
```

Pivot destinations are explicit empty reserved rectangles. The supported initial pivot has one
row dimension and one to four SUM measures. Filter criteria are literal values; existing filter
state is not silently replaced. Validation uses a bounded literal list and refuses existing rules.
Table promotion requires existing headers and a finite bounded source. Protection and object
identity checks stay native. Filter, validation, worksheet lifecycle and layout effects disclose
that reliable automatic compensation is unavailable.

See [dependency navigation](EVIDENCE-DEPENDENCIES.md) for index budgets, incremental SDK calls,
conservative invalidation and unresolved relationship handling.

## Reopen a deterministic task

```ts
const task = {
  version: 1 as const,
  id: 'close-memo',
  surface: 'word' as const,
  stages: [{
    id: 'memo', dependsOn: [],
    program: 'suggest "Previous close" => "Reviewed close"\nfinish when=verified',
  }],
};
for await (const event of session.runDurableProgram(task, {
  approvePlan: reviewExactEffects,
  beforeDispatch: checkCurrentSources,
  revalidate: checkCurrentSourcesAndCompletedOutputs,
})) render(event);

const saved = await session.inspectDurableProgram('close-memo');
// Explicit user action resumes saved.program with NEW validators/approvals, never on startup.
```

The application composition uses IndexedDB in a browser and labels the memory fallback as
non-durable. A saved document, scoped effect journal and program store are required. Each scope
stores at most 32 tasks / 8 MiB; a program is at most 256 KiB. Programs contain bounded command
text needed for continuation; receipts omit inverse bodies, locations, free-form errors and
provenance bodies. No bearer credentials, approvals or executable closures are stored. Current
authentication must match the account that created the session before task reads, resume,
inspection decisions and final dispatch. An account switch requires rebuilding the session.

An interrupted stage needs explicit inspection after stopping other panes. After five minutes,
`reconcileDurableStage` records `unresolved`, `not-applied` or a correlated native `verified`
receipt. Inspection is a trusted operator action; an acknowledgement alone cannot prove a write
landed. A not-applied decision still requires fresh approval on the next attempt. Verified
checkpoints still require current source/output validation before skipping them. A stale pane
whose claim was reconciled cannot later resume its suspended approval and dispatch that claim.

## Coordinate metric handoffs

`composeSession` returns an authenticated `application` port. Register open host sessions with
`OfficeWorkflowCoordinator`, then run an `OfficeHandoffPlan` from `@ge/contracts`. A plan names
the source artifact ID/hash, source binding, destination bindings and at most eight typed steps.
Managed text steps carry a native `TextTargetCapture`, previous generated text and exact metric
row/column bindings. The final Outlook step supplies a subject and a bound text template; it opens
a review form without auto-addressing or sending it.

```ts
const coordinator = new OfficeWorkflowCoordinator();
coordinator.register('source', excel.application);
coordinator.register('memo', word.application);
coordinator.register('deck', powerpoint.application);
coordinator.register('response', outlook.application);
for await (const event of coordinator.run(reviewedHandoffPlan, {
  approvePlan: reviewExactEffects,
})) render(event);
```

This is a compiled SDK integration, not model-authored registration or a new command delimiter.
The source artifact must exist in its originating authorized workspace, be complete, fit 1 MiB
and pass content-hash verification plus live source recapture. The same composition-time Entra
object ID must remain authenticated in every session. Sources and destinations are rechecked
after approval and durable waits. Office has no universal compare-and-swap; a race after the last
native check remains possible. Managed text refuses manual edits. Review-form outcomes stay
`review-required`; an unverified form cannot be replayed as a completed verified stage.

## Outlook checks

The actual send function requires a nonblank subject and at least one recipient. Optional policy
can restrict exact recipient domains and require named non-inline attachments. Capture includes
at most 100 recipients, 50 attachments and 64 KiB of text body. Missing APIs, timeouts, omissions,
item switches or changed facts after registered triggers block the check. These requirements
need the supported compose APIs (Mailbox 1.8 for complete capture); host send-event support is
independently required. A successful body/subject setter followed by stalled readback preserves
the successful mutation with unknown verification. It never invites automatic retry.

Word table reads use `read table:<index>[:<row>:<column>]` with 20×16 cell windows and a 24 KiB
text budget. They disclose omissions and continuations. Ordinary phrase inspection continues
to use content search.

## Release acceptance

```sh
bun run office:lab readiness evidence.json <40-character-commit>
bun run test:native:check native-reports.json <40-character-commit> word,excel,powerpoint,outlook
bun run test:browser-contracts
```

`OfficeReleaseEvidenceSchema` defines the evidence bundle. Readiness exits 0 for `PASSED`, 1 for
`FAIL` and 2 for `BLOCKED_EXTERNAL`. It checks local gates, the advertised host/build/mode matrix,
correlated native fixtures, tenant flagship scenarios and measured outcome targets. Imported
evidence summaries are not cryptographic attestations or deployment grants. No deployment,
tenant provisioning, native Office certification or live-model performance result is implied by
this implementation.

The release target remains at least 95% first-pass success, zero wrong-target/unauthorized-effect/
false-completion regressions and 50% fewer avoidable syncs/repeated-task inference input for the
declared benchmark matrix. These are acceptance targets, not measured product claims.

## Validation recorded for this implementation

The aggregate local suite passes 3,877 tests; 25 optional tests are skipped by default (22 live
scenarios and the three separately enabled browser contracts). All three actual Chromium
IndexedDB tests pass, covering concurrent claims, reopening and rollback. TypeScript, lint/design,
production builds, generated metadata and secret scanning pass. Python passes 76 shared parity
cases, six manifest tests, 40 tooling tests, parser/checker self-tests and 14 deterministic eval
cases; all rebuilt skill archives match their source bytes. Local Bun is 1.3.14; CI uses the
repository's pinned 1.3.13.

Independent security review has no unresolved high or medium findings after account-switch,
stalled Outlook readback, Word inspect fallback and irreversible-preview corrections. Six Tools
panel screenshots were inspected at 320/360/480 across Excel and PowerPoint. Controls wrap without
new horizontal overflow; the older Settings rail label remains clipped. These previews are
scripted browser fixtures, not native Office or tenant evidence.
