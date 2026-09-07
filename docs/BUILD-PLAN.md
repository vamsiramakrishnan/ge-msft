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
