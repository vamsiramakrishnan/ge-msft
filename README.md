# Gemini Enterprise for Microsoft 365

Bring **Gemini Enterprise** into **Word, Excel, PowerPoint, OneNote, Outlook, and Teams** as one
multi-surface Microsoft 365 add-in. The add-in is **client-direct**: it federates the signed-in
user's Entra identity to Google (Workforce Identity Federation, in the browser) and calls Gemini
Enterprise (Discovery Engine `v1alpha`) directly — **no gateway holds credentials, no Google secret
ever reaches a client**. The agent grounds on a composable **research unit** (a NotebookLM notebook
+ federated SharePoint/OneDrive sources + the working document), and every change it makes is
**traceable and reversible** — tracked changes in Word, address-anchored cells in Excel,
citation-tagged blocks elsewhere, each carrying agent id, sources, identity, timestamp, and a
content hash. A surface-agnostic core is written once and reused across thin per-surface bridges.

> **Source of truth.** `CLAUDE.md` is the repo constitution; the **ADRs** (`docs/ADR-000X-*.md`) are
> the current architecture and supersede the older gateway-framed design docs. `docs/STATUS.md` is
> the honest "what's built" inventory. See the [docs index](#docs-index) below.

---

## Architecture

### The layered stack

A surface-agnostic core sits beneath six thin bridges. The core never touches Office.js / TeamsJS /
Graph; the bridges are the only code that does.

```
                         ┌──────────────────────────────────────────────┐
   bridges (host-only)   │ word  excel  powerpoint  onenote  outlook  teams │
                         └───────────────────────┬──────────────────────┘
                                                 │  DocBridge interface
   app core (surface-     ┌───────────────────────┴──────────────────────┐
   agnostic)             │  runtime          web-shell                   │
                         │  AssistSession    NaaAuthClient · composeSession│
                         │  ContextModel     PanelController · React panel │
                         │  Orchestrator     ProvenanceStore · preview     │
                         └───────────────────────┬──────────────────────┘
   capability + I/O       ┌──────────┬───────────┼───────────┬──────────┐
   packages              │ content  │ gemini-    │ graph-    │ triggers │
                         │ blocks/  │ client     │ client    │ events + │
                         │ doc_state│ WIF+assist │ estate    │ the gate │
                         └──────────┴───────────┬┴───────────┴──────────┘
   the boundary                       ┌─────────┴─────────┐
                                      │     contracts     │  types + Zod
                                      │  the grammar, the │  the single
                                      │  manifests, closure│  source of truth
                                      └───────────────────┘
```

- **`contracts`** — the boundary. Types + Zod schemas for the unit descriptor, findings, SSE events,
  provenance, the **capability manifest**, the **CLI command grammar**, the **composition expression
  grammar**, the **skill grammar**, and the **capability-closure** helper. Everything else implements
  against this.
- **`content`** — native-first processing: Office object model → typed `Block`s with host locators →
  token-budgeted, section-aware chunks → the untrusted-wrapped `<doc_state>` snapshot.
- **`gemini-client`** — client-direct Discovery Engine: `WifTokenClient` (Entra→Google token
  exchange), `streamAssist`, `search`, `completeQuery`, `checkGrounding`, `rank`, `SessionContext`.
- **`graph-client`** — Microsoft Graph reader (the estate plane), delegated, as the signed-in user.
- **`triggers`** — the event-driven layer: `HostEvent` lifecycle, `TriggerRegistry`, debounce, and
  the **fail-closed actuation gate**.
- **`runtime`** — the assist loop: `AssistSession` (command loop + composition evaluator + plan
  executor), `ContextModel`, `Orchestrator`, the command compiler.
- **`web-shell`** — the app core plus the React/Vite task pane: `NaaAuthClient` (NAA),
  `composeSession`, `PanelController`, `ProvenanceStore`, host detection, the panel components, and a
  standalone **preview** harness.
- **`bridge-*` / `teams`** — the six surface bridges; each implements one `DocBridge`
  (`getCapabilities` / `listContext` / `resolveContext` / `actuate` / read ports / `watch`).

### Client-direct identity federation (ADR-0001)

```
  Add-in (browser)        Office host         Google STS            Discovery Engine
      │                       │                   │                       │
      │ acquireTokenSilent    │  broker token     │                       │
      ├──────────────────────►│  (Entra OIDC)     │                       │
      │◄──────────────────────┤                   │                       │
      │  STS token exchange (RFC 8693, Entra OIDC → Google access token)  │
      ├──────────────────────────────────────────►│                       │
      │◄──────────────────────────────────────────┤  short-lived token    │
      │  :streamAssist  (Bearer Google token, as the signed-in user)      │
      ├──────────────────────────────────────────────────────────────────►│
      │◄═══════════════ SSE: tokens + citations + provenance ═════════════┤
```

The only optional server piece is a transparent CORS/audit **proxy** (`proxyUrl`) for tenants that
block browser CORS to `discoveryengine.googleapis.com` — a deploy artifact, not a workspace package.
Model Armor, agent routing, and grounding stores are Gemini Enterprise **engine config**, not our
code. The client only ever holds the user's short-lived Entra token and the Google token derived
from it, in memory.

### The capability arc

The product's depth is a four-step arc, each ADR building on the last:

1. **Doc-as-environment (ADR-0003).** The active document is an addressable environment, not a
   payload. An ambient `<doc_state>` snapshot carries structure every turn; the model reads the rest
   **lazily** through narrow, bounded host-read ports (`read`, `search`, `outline`) instead of
   pre-serialising the whole file.
2. **CLI command protocol (ADR-0004).** `streamAssist` has no native function-calling, so the model
   drives the read/write loop by emitting **flat command lines** in a fenced ` ```cmd ` block. The
   runtime parses → validates → **compiles each line into a typed `ActuationRequest`** and runs the
   existing machinery. The command line is the source language; `ActuationRequest` is the IR. A flat
   command verb has no JSON envelope to drift from — empirically more reliable than JSON tool-calls.
3. **Composable algebra + plans + skills (ADR-0005).** A typed **value layer** sits between reads and
   actuations: reads produce `Table`/`Number`/`Text`; pure transforms (`filter`, `select`, `sum`, …)
   compose freely via pipes and `let` bindings; only typed **`Effect`** terminals actuate. *Pure
   composes freely; effects gate* — that single split is both the composition mechanism and the
   safety boundary. The model emits a **plan**, the runtime type-checks it against the manifest,
   **dry-runs** it (reads + pure, zero actuation), previews the effect-set for **one plan-level
   approval**, then executes — each effect gated. **Skills** are named, parameterized compositions:
   the org's capability set grows without shipping code. *(Phases 1–3 implemented;
   `for`/`each` iteration and cross-surface plans deferred — see STATUS.)*
4. **Capability closure (ADR-0006).** A model that composes capabilities over a non-closed set
   composes phantoms. `checkCapabilityClosure` computes `declared manifest ∩ handled kinds ∩ read
   ports ∩ CLI verbs`; per-surface **conformance tests fail the build** on a phantom (an advertised
   capability the bridge can't do). Gaps (handled but not yet reachable by a verb) are tracked on an
   allow-list, not fatal.

### The command surface — `/` verbs and `@` mentions

A surface-agnostic command pane sits on top of the capability stack. The user acts three ways — a
`/` verb, an `@` mention, a prebuilt button — and a right-click does the same from inside the host.
All four compile to the same path; nothing bypasses the gate.

- **`/` verb → an Intent**, scoped per surface by the `CapabilityManifest` (`commandPaletteFor()` in
  `contracts`: `/assist`, `/review`, `/resolve`, `/rewrite`, `/draft`, `/synthesize`, `/notes`). The
  `Composer` opens the verb palette on `/` and the mention picker on `@`, and `parseComposerInput`
  turns a submit into a typed `ComposerInvocation`. A bare question or `/assist` → `send`; any
  actuating verb → the fail-closed `runCommands` plan gate.
- **`@` mention → grounding**, mapped to real `streamAssist` fields — `query.parts[]` (docs/people),
  `toolsSpec.dataStoreSpecs` (connectors), `fileIds` (uploads). Each becomes a removable unit chip.
- **Prebuilt buttons** — `QUICK_ACTIONS` in `contracts` (33 actions, `quickActionsForSurface()`
  closure-filtered) render as the `QuickActionBar`; a `chat` action seeds `send`, a `write`/
  `annotation` action seeds the gate. "Summarize this email", "Review against policy", etc. An action
  with `{{name}}` slots declares typed `parameters` and collects them in a fill form before dispatch —
  a literal placeholder never reaches the model.
- **Context menus** — a right-click "Ask Gemini about this" (`extensions.contextMenus` in the unified
  manifest, `ExtensionPoint` in the OneNote XML) reads the selection and seeds the open pane with it
  as `@this`. The selection rides as data, never instructions; the handoff seed carries no raw text.

The grammar is also carried into Gemini Enterprise as two **skills** (`skill/`), mounted per turn via
`skillsSpec` (`docs/api/discoveryengine/skills-and-agents.md`):

- **`m365-command-planner`** — the front door: turns a free-text `/verb @mentions …` request into a
  structured, confirmable ` ```plan ` block (intent · scope · steps · exclusions · grounding). Its
  `parse_plan.py` is mirrored by `parsePlanBlock()` / `CommandPlan` in `contracts`.
- **`m365-surface-commander`** — the executor: takes the confirmed plan + a live `<doc_state>` and
  emits the ADR-0004 ` ```cmd ` algebra → gate → tracked change / cell / staged draft.

The interaction is mocked in `docs/mockups/6-command-pane.html` (rendered under
`docs/mockups/screenshots/`); the competitive baseline vs Microsoft Copilot is in
`docs/COMPETITIVE-COPILOT.md`.

### The safety spine

Held identically on every surface, enforced at the boundary:

- **Fail-closed actuation gate** — no write executes without explicit approval; `triggers.gate()`
  blocks if no approver is wired. Plans get **one plan-level approval** over a dry-run preview;
  standalone effects get per-write approval.
- **Dry-run before approval** — a plan executes its reads + pure transforms and resolves every effect
  to a Zod-validated `ActuationRequest` **with zero actuation**, so the approval previews exactly what
  will change.
- **Durable provenance** — every actuation carries agent id, sources, identity, timestamp, and a
  content hash. **Word** stamps it into a custom XML part and **Excel** into the workbook settings bag
  (survives save/reopen); the client `ProvenanceStore` view-model lists changes for undo. *(The
  host-metadata write is wired for Word + Excel only; PowerPoint/OneNote/Outlook/Teams persist is not
  yet — see STATUS.)*
- **Untrusted-content boundary** — host document/transcript content is data, never instructions.
  `<doc_state>` and read results are wrapped and framed as data; Model Armor screens at the engine.
  Validated against planted injections in the live-engine probes (ADR-0004 Validation table).
- **`isUnsafeFormula`** — Excel formula writes pass a guard before the gate, so a composed formula
  can't smuggle a dangerous construct into a cell.
- **Content-anchored Word writes** — findings anchor by content (`body.search`), re-resolve at
  apply-time, and degrade to a panel item on drift rather than rendering a broken annotation.

---

## Repo layout

```
packages/
  contracts/        Shared types + Zod schemas — the boundary (unit, finding, SSE, provenance,
                    capability manifest, command/expr/skill grammars, closure helper)
  content/          Native-first content: object model → blocks → budgeted chunks → <doc_state>
  gemini-client/    Client-direct Discovery Engine: WIF exchange, streamAssist, search/rank/grounding
  graph-client/     Microsoft Graph reader (the estate plane), delegated, client-direct
  triggers/         Event layer: HostEvent lifecycle, TriggerRegistry, debounce, the actuation gate
  runtime/          Surface-agnostic core: AssistSession, ContextModel, Orchestrator, command compiler
  web-shell/        App core + React/Vite task pane: NaaAuthClient, composeSession, PanelController,
                    ProvenanceStore, the panel components, and the standalone preview harness
  bridge-word/      Word: native capture + content-anchored tracked changes + watch()
  bridge-excel/     Excel: range capture + address-anchored write-cells + readRange + watch()
  bridge-powerpoint/PowerPoint: slide capture + deck compose + watch()
  bridge-onenote/   OneNote: page synthesis + inline citation tags (web-only, legacy manifest)
  bridge-outlook/   Outlook: mail capture + reviewable reply + the on-send gate
  teams/            Teams: transcript capture + reviewable post-message + meeting events
manifests/          m365-unified.manifest.json (Package A) + onenote.manifest.xml (Package B)
skill/              Gemini Enterprise skill bundles + create/test tooling — the / + @ command
                    surface, carried into the engine: m365-surface-commander (executor) and
                    m365-command-planner (free-text planner)
docs/               ADRs (current architecture), design/contracts/conventions, status, capability
                    map, the API knowledge base, the mockups + rendered screenshots
```

`web-shell` and `runtime` are the bulk of the client and are **surface-agnostic** — they must not
contain Word/Excel/etc.-specific code. Surface specifics live only in `bridge-*` and `teams`.

---

## Commands

```bash
npm install                              # install all workspaces
npm run dev -w packages/web-shell        # run the task-pane dev server (HTTPS; needs a host to sideload)
npm run preview -w packages/web-shell    # see the panel in a plain browser, NO Office host (scripted fixtures)
npm run build                            # build all workspaces
npm run typecheck                        # tsc -b across workspaces
npm run test                             # vitest across workspaces
npm run lint                             # eslint + prettier check
```

`npm run preview` is the fastest way to *see* the product: it mounts the real `<App/>` over a fake
`PanelController` driven by scripted fixtures, with a toolbar to switch surface and toggle every
state (streamed message, citations, context chips, suggestions, run-steps, pending plan, pending
write, proposals, error, busy) — no network, no host, fully clickable.

Copy `.env.example` to `.env` for the engine/tenant config (project, location, engine id, optional
`proxyUrl`) before running against a live engine.

For tenant setup, manifest generation, packaging, Cloudflare dev tunnels, and sideloading choices
across web/desktop hosts, use the [setup guide](setup/README.md).

---

## Status — what's built

Verification baseline: `npm run typecheck` clean · **1538 tests across 128 files green** (Vitest) ·
`npm run lint` clean.

- **All six surface bridges built and tested** — Word, Excel, PowerPoint, OneNote, Outlook, Teams —
  each with an advertised-equals-handled capability set, conformance-enforced per ADR-0006.
- **Client-direct identity** — `NaaAuthClient` (NAA → Entra id + delegated Graph token) and
  `WifTokenClient` (Entra→Google STS exchange, cached, single-flight, epoch-safe). No Google
  credential ever in a client.
- **The capability stack** — doc-as-environment context construction, the CLI command protocol, the
  composable algebra (pure value layer + pipes/bindings, plans + dry-run + plan approval, named
  skills), and capability closure + conformance — all implemented per ADR-0003→0006.
- **Foundational retrieval** — `streamAssist` (grounded, SSE), `search`, `completeQuery`,
  `checkGrounding`, `rank`; the estate read/search path via `graph-client`.
- **The event engine** — per-bridge `watch()` → `HostEvent` → `Orchestrator`, with the fail-closed
  gate handling the rare protective moments (on-send veto, pre-actuation veto).
- **The React/Vite task pane** — the panel components (`App`, `ContextTray`, `MessageThread`,
  `Composer`, `QuickActionBar`, `PlanApprovalCard`, `WriteApprovalCard`, `ProposalCard`,
  `ProvenanceDetail`, `RunSteps`, `SkillsPanel`), MSAL bootstrap, and the standalone preview harness.
- **The `/` + `@` command surface** — `QUICK_ACTIONS` + `CommandPaletteSpec` + `CommandPlan` in
  `contracts`; the `QuickActionBar` and the `Composer` `/`-verb / `@`-mention palettes in `web-shell`;
  right-click context menus in both manifests with a hardened `askSelection` → pane seed. Unit +
  full-stack interplay tested; `security-reviewer` run on the selection-seed flow.

**Partial / deferred** (called out honestly): durable host-metadata provenance writes are wired for
**Word + Excel** but not yet for PowerPoint/OneNote/Outlook/Teams, and the observability/audit sink
over them is undecided; `for`/`each` iteration and cross-surface plans
(ADR-0005 Phase 4); broader CLI verb parity for `slide`/`page`/`mail`/`post` (tracked closure gaps);
estate **writes** (reads are live, Graph/SharePoint writes modeled only); per-capability runtime
detection; the `addContextFile` code-execution upload and the A2UI renderer; **real-host
validation** (the bridges are tested against fakes, not yet sideloaded in a live Office host).

See `docs/STATUS.md` and `docs/CAPABILITY-MAP.md` for the full inventory.

---

## ADR index

The ADRs are the current architecture, in order — each builds on the last.

| ADR | Title | Status | One line |
|---|---|---|---|
| [0001](docs/ADR-0001-client-direct-architecture.md) | Client-direct add-in (no gateway by default) | Accepted | Federate the user's Entra identity to Google (WIF) in the browser and call Discovery Engine directly; no gateway holds credentials. |
| [0002](docs/ADR-0002-capability-model.md) | The capability model (context capture + actuation) | Accepted | Build foundational capabilities — read host objects → context, write actuation ← agent — as the stable layer; experiences compose on top. |
| [0003](docs/ADR-0003-context-construction.md) | Document-as-environment context construction | Accepted | Treat the active doc as a lazily-read, auditably-written environment: an ambient `<doc_state>` + narrow host-read ports, not eager pre-chunking. |
| [0004](docs/ADR-0004-command-protocol-actuation.md) | A command-line protocol for the assist loop | Proposed | The model drives reads/writes via flat CLI command lines; the runtime compiles each to a typed `ActuationRequest`. CLI beats JSON (no envelope drift). |
| [0005](docs/ADR-0005-composable-capabilities.md) | A composable capability algebra | Accepted (Ph 1–3 built) | A typed value layer between reads and effects: pure composes freely, effects gate; plans + dry-run + one approval; skills as saved compositions. |
| [0006](docs/ADR-0006-capability-closure.md) | Capability closure: truthful manifests | Accepted | Compute the closed, executable capability set and enforce it with conformance tests that fail the build on phantom capabilities. |

---

## Docs index

- **Architecture (current):** the six ADRs above — start there.
- `setup/` — tenant prerequisites, Entra/WIF configuration, dev tunnel setup, manifest/package
  generation, sideloading paths, and debugging.
- `docs/STATUS.md` — the honest "what's built / what's deferred" inventory.
- `docs/CAPABILITY-MAP.md` — the per-capability read/write/do inventory across the two planes.
- `docs/CONTRACTS.md` — the authoritative API schemas, the command/expr/skill grammars, closure.
- `docs/CONVENTIONS.md` — stack, code style, testing, security standards.
- `docs/MICROSOFT-ADDIN-CAPABILITIES.md` — the Microsoft 365 add-in platform surface we build on.
- `docs/CONTENT-PROCESSING.md` — the native-first content pipeline in detail.
- `docs/BUILD-PLAN.md` — the original executable checklist *(partly superseded; see its banner)*.
- **Design (legacy, reconciled):** `docs/00-surfaces-plan.md`, `docs/01-architecture.md`,
  `docs/02-design.md`, `docs/03-implementation.md` — the original vision and phasing. These predate
  ADR-0001 and carry **"Superseded / updated by ADR-000X"** banners; read them for the *why*, the
  ADRs for the *what now*.
- `docs/mockups/*.html` — the clickable UX spec, one per surface, plus `6-command-pane.html`
  (the `/` + `@` design); rendered to PNGs under `docs/mockups/screenshots/`.
- `docs/api/discoveryengine/` — the vendored Discovery Engine / Gemini Enterprise API knowledge
  base, including `skills-and-agents.md` (the verified skill create/mount lifecycle).
- `skill/` — the two Gemini Enterprise skill bundles (`m365-surface-commander`,
  `m365-command-planner`) and the create/test tooling, with their own `README.md`.
- `CLAUDE.md` — the repo constitution and how to work here.
