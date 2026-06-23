# Status — what's built (and what's next)

The honest, current state of the codebase. Companion to `BUILD-PLAN.md` (the original checklist,
now partly superseded), `CAPABILITY-MAP.md` (the I/O inventory), and the **ADRs** (the current
architecture). Updated as of the ADR-0006 capability-closure + task-pane wave.

> **Architecture:** client-direct (see `ADR-0001`). The add-in federates the signed-in user's
> Entra identity to Google (Workforce Identity Federation, in the browser) and calls Gemini
> Enterprise (Discovery Engine `v1alpha`) directly. **No gateway by default**; the only optional
> server piece is a transparent CORS/audit proxy via `proxyUrl`. Model Armor, agent routing, and
> grounding are Gemini Enterprise engine config, not our code.

## Verification baseline

`npm run typecheck` clean · **~714 tests across 73 files green** (Vitest) · `npm run lint` clean.

## Packages — built vs planned

| Package | Role | Status |
|---|---|---|
| `@ge/contracts` | Shared types + Zod schemas (the core↔bridge boundary); the command/expr/skill grammars + the closure helper | ✅ built |
| `@ge/content` | Native-first content: object model → blocks → budgeted chunks → `<doc_state>` snapshot | ✅ built |
| `@ge/gemini-client` | Client-direct Discovery Engine: WIF exchange, `streamAssist`, `search`, `completeQuery`, `checkGrounding`, `rank` | ✅ built |
| `@ge/graph-client` | Microsoft Graph reader (Plane B / estate): messages, events, driveItems, users, `/search` | ✅ built |
| `@ge/triggers` | Event-driven layer: `HostEvent` lifecycle, `TriggerRegistry`, debounce, the fail-closed actuation gate | ✅ built |
| `@ge/runtime` | Surface-agnostic core: `AssistSession` (command loop + composition evaluator + plan executor), `ContextModel`, `Orchestrator`, command compiler, skill registry | ✅ built |
| `@ge/web-shell` | App core **+ the React/Vite task pane**: `NaaAuthClient`, `composeSession`, `PanelController`, `ProvenanceStore`, host detection, panel components, the preview harness | ✅ built |
| `@ge/bridge-word` | Word: native capture + content-anchored tracked changes + add-comment + `watch()` | ✅ built |
| `@ge/bridge-excel` | Excel: range capture + address-anchored `write-cells`/`format-cells` + `readRange` + `watch()` | ✅ built |
| `@ge/bridge-powerpoint` | PowerPoint: slide capture + deck compose (`insert-slide`) + slide read + `watch()` | ✅ built |
| `@ge/bridge-onenote` | OneNote: page synthesis + inline citation tags (`append-page`); web-only, legacy manifest | ✅ built |
| `@ge/bridge-outlook` | Outlook: mail capture + reviewable reply + the **on-send gate** | ✅ built |
| `@ge/teams` | Teams: transcript capture + reviewable post-message + meeting events | ✅ built |

**All six surface bridges are built and tested**, each with an advertised-equals-handled capability
set, conformance-enforced per ADR-0006. The `services/gateway` tier was **removed** in the
client-direct reorg — it contradicted ADR-0001.

## The capability stack (ADR-0003 → ADR-0006)

The defining work since the last status: the document became a programmable environment.

- **Doc-as-environment (ADR-0003).** `@ge/content` builds an untrusted-wrapped `<doc_state>`
  snapshot each turn (surface, selection, inventory, named ranges, comments). Each bridge serves
  narrow, bounded, read-only host ports (`captureDocState` / `readRange` / `searchDocument`) so the
  model reads lazily instead of pre-chunking the whole document.
- **CLI command protocol (ADR-0004).** `@ge/runtime` parses the model's fenced ` ```cmd ` block,
  validates each line, and **compiles it to a typed `ActuationRequest`** (`command-protocol.ts`).
  The grammar is scoped per surface by the `CapabilityManifest`. Reads batch freely; writes execute
  one at a time through the gate. Validated reliable + injection-resistant on the live engine
  (ADR-0004 Validation table).
- **Composable algebra + plans + skills (ADR-0005, Phases 1–3 built).**
  - *Phase 1* — a pure value layer (`Table`/`Number`/`Text`) with a transform registry, pipes, and
    `let` bindings; `AssistSession.runCommands` evaluates composed read-expressions
    (`contracts/expr-grammar.ts`, `runtime/compose.ts`).
  - *Phase 2* — a turn's effects form a **plan**: type-check → dry-run (reads + pure, **zero
    actuation**, each effect resolved to a Zod-validated `ActuationRequest`) → `plan-preview` → **one**
    `approvePlan` (fail-closed) → gated execution (`runtime/assist-session.ts`).
  - *Phase 3* — **named skills** (`def name(p…): … end` / call) expand into the same Phase-2 plan;
    bounded substitution, no gate bypass (`contracts/skill-grammar.ts`, `runtime/skill-registry.ts`).
  - *Deferred:* `for`/`each` iteration, durable (host-metadata) skill persistence, and Phase 4
    cross-surface compositions.
- **Capability closure (ADR-0006).** `checkCapabilityClosure({ manifest, handledKinds, readPorts,
  verbKinds })` in `@ge/contracts` computes the closed set; per-surface `capability-closure.test.ts`
  asserts **no phantoms** (hard) and tracks **gaps** against an allow-list. Drift remediation
  shipped: phantom actuations un-advertised, Excel `readRange` added, `reply → comment-reply` verb
  exposed. Runtime requirement-set *detection* in `getCapabilities()` remains a follow-up — manifests
  are static but now conformance-checked.

## Other capabilities delivered

- **Identity, client-direct.** `NaaAuthClient` (MSAL Nested App Auth) yields the Entra id token (WIF
  subject), delegated Graph tokens (Plane B), and the identity for provenance. `WifTokenClient`
  exchanges Entra→Google (RFC 8693) in the browser, cached with TTL/skew, single-flight, epoch-safe
  invalidation. No Google credential ever reaches a client.
- **Grounded assist loop.** `AssistSession` ties a bridge to `streamAssist`: attach context → stream
  a grounded answer (tokens + citations + provenance) → run the command/plan loop → apply reversible,
  provenanced actuation. Session id is captured and resumable across surfaces.
- **Foundational retrieval (beyond assist).** `search` (faceted/filtered, boost, snippets,
  pagination, `dataStoreSpecs`), `completeQuery` (type-ahead), `checkGrounding` (per-claim score —
  backs the on-send / pre-actuation gate, fail-closed), `rank` (semantic rerank).
- **Native-first content processing.** Office object model → typed `Block`s with host locators →
  token-budgeted, section-aware chunks → contextualized → `ResolvedContext` mapping 1:1 to
  `query.parts`. Budget picks inline / reference / upload-for-code-execution.
- **Two context planes.** Plane A (in-document, Office.js, per bridge) and Plane B (estate, Graph:
  search SharePoint/OneDrive/mail/calendar/people as the user).
- **Event-driven, not assistant-spamming.** `watch()` on each bridge emits `HostEvent`s; the
  `Orchestrator` debounces and routes. Most events **construct context** via the `ContextModel`
  (cheap, no model call); the `TriggerRegistry` gate handles the rare protective moments (on-send
  grounding veto, pre-actuation veto); suggestions are scarce and ignorable.
- **Reversible, provenanced writes.** Every actuation carries agent id, sources, identity, timestamp,
  and a content hash; Word anchors by content (`body.search`, re-resolved at apply-time, degrades on
  drift); Excel formula writes pass `isUnsafeFormula`; Outlook/Teams open reviewable forms rather
  than sending silently. Word + Excel persist a durable provenance record into host metadata (custom
  XML part / workbook settings) after a write lands. `PanelController` stages proposals/plans for
  explicit confirmation and records outcomes in the client `ProvenanceStore` view-model.
- **The React/Vite task pane.** The panel renders the real components (`App`, `ContextTray`,
  `MessageThread`, `Composer`, `PlanApprovalCard`, `WriteApprovalCard`, `ProposalCard`,
  `ProvenanceDetail`, `RunSteps`, `SkillsPanel`) over `PanelController`; MSAL bootstrap +
  bridge-selection wiring; **`npm run preview -w packages/web-shell`** mounts `<App/>` over scripted
  fixtures with no Office host.

## Testing approach

Vitest across all workspaces (~714 tests / 73 files). Bridges are tested against **in-repo Office
fakes** (`web-shell/src/test-harness/fake-{office,word,excel,powerpoint}.ts`), not a live host.
Coverage includes: contract schema round-trips; the command/expr/skill grammars; per-surface
**capability-closure conformance** (no phantoms; gaps within the allow-list); capture + actuate
plans; the composition evaluator + plan dry-run/execute; the orchestrator + gate; property tests for
chunking, JSON-stream parsing, and Word actuate plans; and task-pane integration tests
(`web-shell/src/taskpane/*-integration.test.ts`) that drive the panel against the fakes.

## What's next (pending)

1. **Real-host validation** — the bridges pass against fakes; sideloading the manifest in a live
   Office/Teams host and validating each surface against its mockup is not yet done.
2. **Durable provenance writes** — wired for **Word** (a custom XML part via `customXmlParts.add`,
   gated on WordApi 1.4) and **Excel** (the workbook settings bag), so writes stay provenanced across
   save/reopen; the pure serializer is `bridge-{word,excel}/src/provenance-record.ts`. **Not yet wired
   for PowerPoint/OneNote/Outlook/Teams**, and no observability/audit sink aggregates the records — so
   the cross-session trail is partial.
3. **Broader CLI verb parity** — `slide`/`page`/`mail`/`post` verbs and other surface read ports are
   tracked **gaps** on the closure allow-list (handled by the bridge, not yet reachable by a verb).
4. **ADR-0005 Phase 4** — `for`/`each` iteration, durable skill persistence, cross-surface plans
   (read Excel → write PowerPoint).
5. **Estate writes** — Graph/SharePoint reads are live; writes (send mail, create event, upload /
   checkout) are modeled but not wired.
6. **Security hardenings noted by review** — keep `decideSend` total; bound the on-send trigger with
   a timeout so a *hung* trigger cannot wedge Send; validate `location`/`proxyUrl` at config
   construction.
7. **Engine extras** — the v1 `addContextFile` path (code-execution uploads) and the A2UI renderer —
   designed, not built.

See `MICROSOFT-ADDIN-CAPABILITIES.md` for the Microsoft 365 add-in surface we build on, and
`CAPABILITY-MAP.md` for the per-capability read/write inventory.
