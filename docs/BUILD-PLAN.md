# Build Plan

> **⚠️ Superseded in part — read `STATUS.md` for the current state.** This plan was written for the
> original **gateway-based** three-tier design. The project has since pivoted to **client-direct**
> (`ADR-0001`): no gateway, identity federated in the browser (WIF), Gemini Enterprise called
> directly, agents/Model Armor are engine config. The gateway tasks below (0.3–0.9, 1.2, the
> `services/*` work) are **obsolete** and intentionally not done; their intent is delivered
> client-side instead. What *is* built — the surface-agnostic core, the foundational retrieval
> clients, the Word/Excel/Outlook/Teams bridges, the event engine, and the web-shell core — is
> inventoried in `STATUS.md` and `CAPABILITY-MAP.md`. Treat those two as the source of truth for
> status; treat the per-surface ACs and mockups below as still-useful acceptance targets.

The executable checklist for this project. Work top to bottom, one task at a time. Use `/plan` at the start of each phase. After finishing a task: run typecheck + tests + lint, confirm the acceptance criteria, run `security-reviewer` if it touched auth/credentials/guardrails/provenance, then change its `[ ]` to `[x]` and commit.

Each task lists the package(s) it touches and a verifiable acceptance criterion (AC).

---

## Phase 0 — Foundation (retire the hard risks first)

The goal of P0 is the slice-1 spine: a signed-in user, federated to Google, gets a grounded streamed answer over a Word selection. If that works, identity federation and the streaming relay — the two hardest pieces — are solved.

- [x] **0.1 Monorepo scaffold.** Root `package.json` (Bun workspaces), `tsconfig.base.json`, ESLint + Prettier, Vitest config, the `packages/*` and `services/*` directories with placeholder `package.json` each. *AC: `bun install` succeeds; `bun run typecheck` runs clean on empty packages.*
- [x] **0.2 Contracts package.** Implement `packages/contracts` exactly per `docs/CONTRACTS.md`: the `Intent` enum, `UnitDescriptor`, `AssistRequest`, `Finding`, `ProvenancePayload`, and the SSE event types, as TypeScript types + Zod schemas. *(Done, and since grown well beyond the original scope: the capability manifest, the command/expr/skill grammars, and the closure helper — see ADR-0002→0006.)* *AC: schemas parse the example payloads in CONTRACTS.md; exported types compile.*
- [ ] **0.3 Gateway skeleton.** `services/gateway` Fastify app with `/healthz`, structured logging, config loaded from env (`.env.example`), and CORS for the web-shell origin. *AC: `GET /healthz` returns 200 with build info.*
- [ ] **0.4 Entra token validation.** Middleware that validates the inbound Entra JWT against the tenant JWKS (issuer, audience, expiry) and extracts identity + `groups`/`roles` claims. *AC: a valid test token passes; tampered/expired tokens are rejected with 401. (run `security-reviewer`)*
- [ ] **0.5 Identity federation.** Exchange the validated Entra OIDC token for scoped Google credentials via Workforce Identity Federation (STS). Cache the federated token under its TTL. *AC: a federated token is obtained and reused; expiry triggers re-exchange. (run `security-reviewer`)*
- [ ] **0.6 StreamAssist client + SSE relay.** Call the Gemini Enterprise `streamAssist` method using the federated credentials and re-emit the server stream as SSE per the CONTRACTS event format, appending grounding citations. Polling fallback behind the same interface. *AC: `POST /assist` streams tokens + citations as SSE for a fixed test query.*
- [ ] **0.7 Unit resolver.** Resolve a `UnitDescriptor` server-side: federated connector sources fetched live with the user's identity; notebook reference attached; surface context included. *AC: a descriptor with a notebook + one federated SharePoint source resolves into a grounding scope the assist call uses. (run `security-reviewer`)*
- [ ] **0.8 Model Armor screening.** Screen incoming host content (untrusted) before it reaches a model; screen outputs. *AC: a prompt-injection test string in the doc context is flagged and neutralized. (run `security-reviewer`)*
- [ ] **0.9 Provenance signing.** Attach `ProvenancePayload` (agent id, sources, identity, timestamp, content hash) to responses. *AC: every assist response carries a well-formed, hashable provenance payload.*
- [ ] **0.10 Web-shell shell.** `packages/web-shell`: React panel, `UnitComposer` (notebook + connector chips), `AuthClient` (MSAL NAA with Office Dialog fallback), `StreamClient` (SSE consumer + polling fallback), `ProvenanceStore` interface. Surface-agnostic. *AC: the panel renders, authenticates a user via NAA, and streams an answer from the gateway.*
- [ ] **0.11 Slice-1 end to end.** Wire a minimal Word task pane to the web-shell: select text → `/assist` → grounded streamed answer in the panel. *AC: the full chain works against a real Gemini Enterprise app + Entra tenant. This is the P0 exit gate.*

## Phase 1 — Word (the deepest surface, the bellwether)

> **Bridge status:** the Word bridge (capture + content-anchored tracked changes + add-comment +
> `watch()`) is **built and tested against in-repo Office fakes** — see `STATUS.md`. The per-surface
> ACs below remain the targets for **real-host** validation, which is still pending. The intent-named
> A2A flows (1.2/1.4/1.5) are reframed by `ADR-0004`/`ADR-0005`: the assist loop is the CLI command
> protocol + composable plans, not intent dispatch.

- [ ] **1.1 Word DocBridge.** `packages/bridge-word`: read selection/body/content controls/`getFileAsync`; write tracked changes (`changeTrackingMode`), `insertOoxml`, comment replies. *AC: round-trip read selection → write a tracked change.*
- [ ] **1.2 Review agent (A2A).** `services/agents`: an ADK Review agent exposed as an A2A server, deployed to Agent Engine, returning `Finding[]` per CONTRACTS. Gateway routes intent `review` to it. *AC: `POST /review` returns structured findings grounded on the unit.*
- [ ] **1.3 Annotations + anchoring.** Render findings as inline annotations (annotations API) anchored by `matchText`/`contextHint` via `body.search`; re-resolve at apply-time; degrade to a panel item if the text is gone. *AC: clicking a finding shows its card; Accept applies a tracked change at the right range; an edited-away finding degrades gracefully.*
- [ ] **1.4 Comment task queue.** On comment events, route intent `resolve-comment`; apply the edit as a tracked change, post a threaded reply, resolve the comment. *AC: a reviewer comment is resolved end to end with an edit + reply.*
- [ ] **1.5 Surgical regeneration.** Scope intent `regen-clause` to a content control; stream the rewrite into it while locking others client-side. *AC: only the targeted clause changes.*
- [ ] **1.6 Provenance (custom XML).** Implement `ProvenanceStore` for Word: write/read the custom XML part (agent, sources, hash, identity, timestamp, StreamAssist session id). *AC: provenance survives save/reopen; the session resumes.*
- [ ] **1.7 Word exit gate.** Run a full FSI contract-review workflow against a design-partner doc. *AC: the §5 "matches the mockup" walkthrough (`docs/mockups/1-word.html`) works on a real document.*

## Phase 2 — Excel + PowerPoint (reuse the spine; surface code only)

> **Bridge status:** both the Excel bridge (range capture + `write-cells`/`format-cells` + `readRange`
> + `watch()`) and the PowerPoint bridge (slide capture + deck compose + `watch()`) are **built and
> tested against fakes**. The `=GE.ASK` streaming custom function is now ALSO implemented
> client-direct (`packages/web-shell/src/functions/`, wired via `functions.json` + the Excel XML
> manifest's `FunctionFile`/`ExtendedOverrides`) — real-host activation still pending. PowerPoint
> speaker-note generation stays out of the bridge: Office.js ships no notes write path (see 2.3).
> Real-host validation pending.

- [~] **2.1 Excel streaming function.** `=GE.ASK(prompt, range)` as a streaming custom function,
  client-direct: `packages/web-shell/src/functions/` (budgeted untrusted-data framing +
  `CustomFunctions.associate` seam + SSE→cell streaming adapter), `public/functions.json`
  metadata, vite `functions.html` entry, and the Excel dev XML manifest FunctionFile +
  ExtendedOverrides. *AC: a cell streams a grounded answer — code complete and unit-tested;
  REMAINING: activate against a real Excel host + engine.*
- [ ] **2.2 Excel entity cells.** Linked-entity load service backed by the gateway/Gemini retrieval; vendor cells expand into agent-enriched cards. *AC: an entity cell loads on demand and expands; nothing large is stored in the workbook.*
- [~] **2.3 PowerPoint composer.** Deck compose + shape-text writes landed; provenance via the shared store. **Speaker notes remain platform-blocked**: no Office.js write path exists (PowerPointApi through 1.10; OfficeDev/office-js#3269 open) — the bridge deliberately advertises nothing it cannot actuate (ADR-0006). *AC: slides portion matches `docs/mockups/3-powerpoint.html`; notes portion waits on the host API.*
- [ ] **2.4 P2 exit gate.** An Excel analyst flow and a PowerPoint deck flow both ground on the *same* unit a Word user assembled. *AC: unit continuity demonstrated across three surfaces.*

## Phase 3 — OneNote + Teams (divergent client models)

> **Bridge status:** the OneNote bridge (page synthesis + inline citation tags), the Teams bridge
> (transcript capture + reviewable post-message + meeting events), and the Outlook bridge (not in the
> original five — mail capture + reviewable reply + the on-send gate) are all **built and tested
> against fakes**. The Bot Framework bot + message extension (3.4) are not built. Real-host
> validation pending.

- [ ] **3.1 OneNote package.** `packages/bridge-onenote` with its own legacy XML manifest (`manifests/onenote.manifest.xml`); `OneNote.run` page synthesis. *AC: the add-in loads in OneNote on the web.*
- [~] **3.2 OneNote research capture.** Citation-tagged page synthesis is landed in
  `packages/bridge-onenote`. **NotebookLM audio/video overview remains blocked: NotebookLM exposes
  no public API** to call from the add-in; keep this `[~]` until one exists or a supported export
  path is chosen. *AC: synthesis portion done in-bridge; overview waits on an API.*
- [ ] **3.3 Teams meeting app.** `packages/teams`: host the web-shell as a meeting side panel via TeamsJS; ground on the unit + live transcript (RSC consent); intent `meeting-notes` → live notes + grounded action items. *AC: live notes + action items appear in a Teams meeting.*
- [ ] **3.4 Teams bot + message extension.** Bot Framework bot for "ask the agent" rendering Adaptive Cards; a message extension to ground a message on the unit. *AC: the bot answers grounded; the message extension returns a grounded result; recap card posts to a channel.*
- [ ] **3.5 P3 exit gate.** The cross-surface flow (OneNote → Word → Excel → PowerPoint → Teams) runs unbroken across at least three surfaces with one unit + one identity. *AC: continuity demo recorded.*

## Phase 4 — Continuity, hardening, distribution

- [ ] **4.1 Cross-surface continuity.** The unit and provenance trail persist and resume across surfaces and sessions. *AC: a unit assembled in OneNote is reused in Teams without re-establishing.*
- [ ] **4.2 Residency pinning.** Per-tenant region config for Cloud Run + `discoveryengine`; VPC firewall to connector FQDNs. *AC: a configured tenant keeps data in-region end to end. (run `security-reviewer`)*
- [ ] **4.3 Audit + observability.** Structured invocation log → BigQuery; cross-surface tracing; latency budgets. *AC: every invocation (who/agent/sources/changes) is queryable.*
- [ ] **4.4 Packaging + distribution.** Build the unified M365 package (A) and the OneNote package (B); validate with the Agents Toolkit; prepare AppSource / Teams Store / admin-center deployment. *AC: both packages validate and sideload cleanly across declared platforms.*
- [ ] **4.5 Admin governance.** Consent flows, Package Management API controls, Teams admin controls. *AC: an admin can enable/disable per tenant.*

---

### Notes for the implementer
- If a task is blocked on a real GCP/Entra/Gemini resource that isn't provisioned, implement against the contract with a clearly-labelled fake/mock, mark the task `[~]` (in progress), and note the dependency — don't fake silently.
- Keep `packages/web-shell` free of surface-specific code. If something feels surface-specific, it belongs in a `bridge-*`.
- Re-read `docs/CONTRACTS.md` before changing any cross-boundary type.

## Workspace UX upgrade — September 2026

This work extends the implemented client-direct architecture; the historical gateway tasks above
remain historical release targets rather than prerequisites for this UI work.

- [x] Visible attach/remove/reveal context chips and capability-filtered starting tasks.
- [x] Searchable 47-action library with output filters and browser-local catalog pins.
- [x] Typed intent controls, request source chips, response options, and IME-safe submission.
- [x] Twelve outcome workflows using existing bridges and approval authorities.
- [x] Completed-answer reuse, streaming insertion restrictions, and proposal concurrency guard.
- [x] Change/target summaries and exact command/dry-run inspection in approval.
- [x] Interactive sample preview with the real controller and six scripted host scenarios.
- [x] Validation: types, lint, 2,149 passing tests (10 skipped), production build, language/resource
      drift checks, and narrow-pane browser inspection. Browser checks cover source chips, action
      search/pins, all six host entry points, completed answers, and the real controller's sample
      approval flow at 320/360 px widths and 480/760 px heights.
- [ ] Live Office sideloading against a configured Entra/Gemini tenant (separate release gate).


## Runtime extension foundation — September 2026 (ADR-0011)

User-directed foundation work builds on the existing client-direct runtime and the workspace upgrade.

- [x] Typed lifecycle phases for request receipt, model calls/events/responses, tools, plans, effects,
      outcome verification, and completion, wired through the real AssistSession entry points.
- [x] Ordered guard/observer dispatch, deadlines, abort signals, immutable isolated inputs, bounded
      provider context, metadata diagnostics, and atomic namespaced extension installation/disposal.
- [x] Production host-event ownership, independent debouncers, bounded serialized ingress, cancellation
      and late-result suppression, idle context refresh, and explicit suggestion actions.
- [x] Same extension definition bundle in the pane and separate Outlook send runtime; required trigger
      check failures block, while post-operation observer failures preserve landed outcomes.
- [x] Task outcome records and default incomplete-execution verifier; retain supplied inverse receipts
      without automatic retry or rollback. Preserve existing human approval and delegated identity.
- [x] Developer API guide and ADR with request/context/event/verification examples and actual boundaries.
- [x] Final verification: typecheck, 2,187 passing tests (10 skipped), lint, production build,
      and language/resource drift checks. Reviewed approval, cancellation, and provenance paths.
- [ ] Live host validation against a configured Entra/Gemini tenant (separate release gate).

## Bounded programs and isolated command execution — September 2026 (ADR-0013)

- [x] Typed artifact bindings and a versioned SDK program compile through the existing command path.
- [x] Macro-aware preflight budgets, complete framing and runtime-verified completion; failures and
      cancellation never imply successful completion or trigger automatic replay.
- [x] Compact command discovery, targeted help, shorter skill bootstrap and bounded result references
      with projected inspection; TypeScript/Python parser and preflight parity maintained.
- [x] Direct Discovery Engine v1alpha sessionless option, verified against the current API reference.
      Command/planner requests use complete bounded context by default; normal chat remains stateful.
- [x] Independent requests retain task/program/result continuity, fresh snapshots and active grounding.
      Overflow and session-bound uploads stop explicitly. Adapter mutation cannot leak chat session
      IDs into observers, writes or task state.
- [x] Real DuckDB acceptance: four model calls become one (zero via typed SDK), with one approval.
      Full/compact and conversation/sessionless controls distinguish round-trip savings from replay.
- [x] Security review closed for completion, approval, macro expansion, result limits and session isolation.
- [x] Final release checks: types, 2,408 passing tests (10 skipped), lint, production build and
      language/resource drift checks.
- [ ] Live tenant check: sessionless private-skill routing, saved conversation history, p50/p95 latency,
      Office application/readback, and deployed worker/CSP behavior remain release gates.

## Workflow recipes and execution-state UX — September 2026

- [x] Normalize typed Excel grids consistently across planning, application, readback and recovery.
- [x] Project deterministic execution state for independent model requests; retain full inspectable
      records, unresolved errors, bindings and uncertain effects without replaying the transcript.
- [x] Add three typed, versioned recipes with parameter schemas, constrained computation, conditional
      materialization and the existing approval/freshness/verification authority.
- [x] Provide workflow cards, guided parameters, explicit saved settings, result previews and truthful
      write/recovery outcomes in the narrow pane, without model calls for known recipes.
- [x] Extend real-WASM controls and live API evaluation with session-mode and latency reporting.
- [x] Complete UI/state bugbash, security review, types, tests, lint, build and generated-file checks.
      React .tsx tests now run in CI. Local suite: 2,542 passed, 16 optional live cases skipped;
      security review closed on 148 targeted regressions. See ADR-0014 and WORKFLOW-RECIPES.
- [ ] Run real Office and tenant acceptance, including narrow-pane browser layout, sessionless
      private-skill routing, saved-history behavior and measured p50/p95. The cloud browser could not
      reach the local preview; no live provider or Office performance result is claimed.

## Architecture unification — September 2026

- [x] Centralize request admission and correlated effect receipts for every execution route.
- [x] Derive bridge dispatch coverage from executable handlers; share provenance primitives while
      keeping host persistence and mutation boundaries explicit.
- [x] Move analysis/program schemas to contracts with compatibility exports; remove duplicated
      producer classifications and align generated skill/CLI metadata with executable grammar.
- [x] Consolidate command context lifecycle and panel execution ownership; preserve public behavior,
      cancellation, queue routing, provenance and exact approval semantics.
- [x] Fail closed on missing/corrupt CLI contracts and remove unsupported planner disclosures.
- [x] Reconcile architectural documentation with the client-direct implementation and add durable
      conformance checks for boundaries that previously drifted.
- [x] Complete security review, type/lint/build gates, existing behavior tests and CLI/skill parity;
      publish the refactor with measured structural changes and explicit remaining limitations.

Validation: 2,671 Vitest tests passed, 16 optional live tests skipped (198 files); all workspace
builds, TypeScript, ESLint, Prettier, frozen lockfile, generated-resource checks, 46 Python tests,
71 parity cases, 14 deterministic eval cases, and three source-exact skill archives passed.
Security review closed all concrete findings. Live Office/tenant acceptance remains tracked above.


## User-directed bounded document context — October 2026

- [x] Separate metadata snapshots from bounded content pages across Word, Excel and PowerPoint.
- [x] Traverse complete sources with explicit cursors, EOF, extraction scope and limitations.
- [x] Summarize with bounded sessionless map/reduce requests and verify a second source traversal.
- [x] Fix natural-language `read` routing and stop using an outline as whole-document content.
- [x] Add scale, request-isolation, drift, cancellation and incomplete-response regressions.
- [x] Complete type, test, lint, build and generated-skill parity gates; review changed trust boundaries.
- [ ] Live Office and tenant acceptance: pagination, requirement sets, sessionless routing, latency
      and semantic summary quality. Mock coverage is not a live performance result.

Implementation and operating limits: [DOCUMENT-CONTEXT.md](DOCUMENT-CONTEXT.md).

Validation: 2,692 Vitest tests passed, 16 optional live tests skipped (202 files); TypeScript,
ESLint, Prettier, all workspace production builds, generated-resource/manifest drift checks,
46 Python tests, 71 parity cases, 14 deterministic eval cases and source-exact skill archives passed.
Reviewed source-data framing, hook guards, policy propagation, session isolation, cancellation,
coverage/freshness claims and absence of new write/auth/credential paths. Live acceptance remains
open above.

## User-directed reusable evidence infrastructure — October 2026

- [x] Extract source traversal, coverage and freshness validation from document summarization.
- [x] Add immutable, content-addressed snapshots, a bounded storage port and exact retrieval.
- [x] Add task-independent context projections with required evidence, structural expansion and explicit omissions.
- [x] Expose snapshots through DocFs and bind optional evidence dependencies to command/SDK execution.
- [x] Exercise atomic capture, isolation, byte budgets, reuse, stale evidence and write-scope enforcement.
- [x] Complete repository gates and publish the shared design and extension contract.

Validation: full-suite run passed 2,705 tests with 16 optional live skips (203 files); the final
navigation/Unicode-window and execution refinements passed all 53 targeted tests, including two
additional cases. TypeScript, ESLint, formatting, production builds, generated manifest/resource
checks and source-exact skill bundle validation passed. Reviewed immutable publication, storage
integrity, principal isolation requirements, untrusted data framing, read/write scope separation,
post-approval freshness, cancellation, and uncertain-write invalidation. Live host acceptance remains
open. Storage is in-memory by default; automatic graph extraction and multi-write rebasing are not
claimed. See DOCUMENT-CONTEXT.md for the complete extension contract and limits.

## User-directed evidence skill alignment — October 2026

- [x] Route absolute-path CLI cat/grep through DocFs while preserving named workspace artifacts.
- [x] Teach commander bounded evidence traversal, continuations, coverage and binding freshness.
- [x] Teach planner to preserve selected edit targets while describing broader evidence reads.
- [x] Document reproducible skill packaging and the separate live provisioning step.
- [x] Regenerate resources/manifests, validate behavior and publish the skill update.

Validation: 2,709 Vitest tests passed, 16 optional live tests skipped (203 files). TypeScript,
ESLint, formatting, production builds, generated manifest/resource drift checks, 46 Python tests,
71 parity cases, 14 deterministic eval cases and all three source-exact skill archives passed.
An isolated skill forward-test selected the supplied text continuation before proposing a redline.
Commander bootstrap remains below four KiB. Reviewed DocFs-only absolute path routing, hook guards,
source-data trust, native target separation and evidence freshness. Tenant provisioning is separate
and was not performed by this repository update.

## User-directed PR 13 review fixes — October 2026

- [x] Bound Excel snapshot value extraction before reading headers; retain addresses and dimensions.
- [x] Preserve worksheet/named-range omissions and PowerPoint total count/truncated inventories.
- [x] Preserve Word selection token boundaries while supporting long and multi-paragraph edits.
- [x] Bind same-batch cell rebasing to the actual correlated, verified output hash.
- [x] Use paged automatic document fallback and bounded serialized chat context with explicit coverage.
- [x] Complete repository checks and independent security review before merging the fixes.
- [ ] Live Office/Gemini tenant acceptance of final-head large-source reads and coauthoring races.

The regressions use labelled Office simulators; they do not establish live tenant or host performance.

Validation: frozen Bun 1.3.13 install, TypeScript, ESLint/Prettier/design checks, all production
builds and generated skill/resource checks passed. Vitest passed 3,663 tests with 22 optional live
skips (247 files). Skill CI passed 74 corpus cases, six manifest tests, 40 tooling tests and
14 deterministic eval cases, including rebuilt exact-source archives. Independent security review
covered native-target freshness, selection authority, serialized context bounds and Python framing.
The Excel simulator now consumes queued writes once and preserves formula identity, so a verified
formula hash matches a fresh capture without relaxing the production freshness check.


## User-directed verified Office foundation — October 25, 2026

This implementation begins the 100x capability plan; the full roadmap is not complete.
See `VERIFIED-OFFICE-WORKFLOWS.md` for exact implemented and remaining boundaries.

- [x] Negotiate native host support and recheck effective capabilities before dispatch.
- [x] Add initial strict capability modules and descriptive scaffolding.
- [x] Authenticate indexed evidence pages and bound incremental navigation.
- [x] Batch independent Excel reads, bound formula footprints and wait for calculation state.
- [x] Capture Word/PPT native text targets and preserve manual edits during managed refresh.
- [x] Add transactional durable effect claims without replaying approvals or uncertain writes.
- [x] Add contract-review/response entry points through the existing executor.
- [x] Extend deck compiler with editable charts, themes and explicit layout checks.
- [x] Add native fixtures, exact-commit report checking and host doctor tooling.
- [x] Complete browser IndexedDB and 320/360/480 workflow panel screenshot acceptance.
- [ ] Run final-commit native Office fixtures and tenant model scenarios.
- [x] Expand strict modules across every installed Word/Excel/PowerPoint/Outlook handler; retain truthful native verification limits.
- [x] Add bounded captured formula/name/table dependencies, reverse impact and unchanged-page reuse.
- [x] Add durable deterministic stages, explicit inspected recovery and authenticated cross-application SDK coordination.
- [x] Add bounded native Excel operations, richer Word reads, Outlook send checks, template/layout capture and managed artifact/chart refresh.
- [ ] Certify supported operations and flagship workflows on the advertised native host/build matrix; record exact-head tenant outcomes and benchmark targets.

See `OFFICE-PHASES.md` for the supported code paths, usage and explicit refusals. In-place native
PowerPoint chart-series editing, full template asset cloning, dynamic/external Excel dependencies
and arbitrary process-to-process Office transport are not advertised. Broader OneNote/Teams and
estate operations remain outside the first four-surface certified core.

Security review identified external workbook formula aliases, aggregate fallback limits and
tracking-mode cleanup receipt loss; regression fixes are included. Native/browser acceptance
is recorded separately and must not be inferred from the simulator suite.

Browser acceptance: two real Chromium IndexedDB tests pass (cross-pane claims/reopen and
transaction rollback). Inspected 21 screenshots at 320/360/480: filled contract review,
managed text refresh, response package, native fixture controls, unavailable report and
interrupted changes before/after inspection acknowledgement. New controls do not overflow;
pre-existing Settings rail text clipping remains outside this increment. CI retains screenshots.
