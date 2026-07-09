# CLAUDE.md — Gemini Enterprise for Microsoft 365

This file is the source of truth for how this repo works. Read it fully before doing anything.

## What we're building

A multi-surface Microsoft 365 add-in that brings **Gemini Enterprise** into Word, Excel, PowerPoint, OneNote, and Teams. It is **client-direct** (ADR-0001): the add-in federates the **signed-in user's identity end to end** to Google (Workforce Identity Federation) and calls Gemini Enterprise directly — no gateway holds credentials, no Google secret ever reaches a client. A surface-agnostic core (`runtime` + `web-shell`) is reused across thin per-surface bridges. The agent grounds on a composable **research unit** (a NotebookLM notebook + federated SharePoint/OneDrive sources + the working document), and leaves changes that are **traceable and reversible**.

The full reasoning lives in `docs/` — read these before implementing the corresponding layer:
- `docs/01-architecture.md` — the three tiers, the gateway internals, identity federation, the anchoring contract.
- `docs/02-design.md` — the five experience invariants, the per-surface verbs, the build phases.
- `docs/03-implementation.md` — packaging, per-surface APIs, repo layout.
- `docs/CONTRACTS.md` — **authoritative** API schemas, the unit descriptor, the finding schema, the SSE protocol. Implement against this exactly.
- `docs/CONVENTIONS.md` — stack, code style, testing, security standards.
- `docs/BUILD-PLAN.md` — the **executable checklist**. Drive all work from this.
- `docs/mockups/*.html` — the UX spec. Open these to see the intended interaction for each surface before building that surface's client.

## How to work in this repo

1. **Drive everything from `docs/BUILD-PLAN.md`.** Implement one task at a time, top to bottom. Don't skip ahead.
2. **Use plan mode (`/plan`) at the start of each phase** to scope it before writing code. Use the built-in Explore subagent to scan the repo so the main context stays clean.
3. **Implement against `docs/CONTRACTS.md`.** The shared types in `packages/contracts` are the boundary between gateway and clients — change them deliberately and update both sides.
4. **After any task that touches identity, credentials, guardrails, or provenance, invoke the `security-reviewer` subagent** before marking the task done.
5. **A task is done only when** types pass, tests pass, lint is clean, and the task's acceptance criteria in the build plan are met. Then check it off in `docs/BUILD-PLAN.md`.
6. Use `/next-task` to pick up the next unchecked item automatically.

## Repo structure

**Architecture note:** we are **client-direct** (see `docs/ADR-0001-client-direct-architecture.md`).
The add-in federates the signed-in user's Entra identity to Google (Workforce Identity Federation,
browser-side) and calls Gemini Enterprise (Discovery Engine) directly — there is **no gateway by
default**. The only optional server piece is a transparent CORS/audit proxy, configured via
`proxyUrl`; it is a deploy artifact, not a workspace package. Model Armor, agent routing, and
grounding are Gemini Enterprise engine config, not our code.

```
packages/
  contracts/        Shared TypeScript types + Zod schemas (the core↔client boundary; no gateway)
  content/          Native-first content processing: host object model → blocks → budgeted chunks
  gemini-client/    Client-direct Discovery Engine: WIF token exchange, streamAssist, search/rank/grounding
  graph-client/     Microsoft Graph reader (Plane B / estate), delegated, client-direct
  triggers/         Event-driven layer: HostEvent lifecycle, TriggerRegistry, debounce, the actuation gate
  runtime/          Surface-agnostic core: DocBridge/AuthClient, AssistSession, ContextModel, Orchestrator
  web-shell/        The reused web app core: AuthClient (NAA), composeSession, PanelController, ProvenanceStore
  bridge-word/      Word DocBridge: native capture + content-anchored tracked changes + watch()
  bridge-excel/     Excel DocBridge: range capture + address-anchored write-cells + watch()
  bridge-outlook/   Outlook DocBridge: mail capture + reviewable reply + the on-send gate
  bridge-powerpoint/PowerPoint: slide capture + deck composer + watch()
  bridge-onenote/   OneNote page synthesis + inline citation tags (web-only, legacy manifest)
  teams/            Teams DocBridge: transcript capture + reviewable post-message + meeting events
manifests/          m365-unified.manifest.json (Package A) + onenote.manifest.xml (Package B)
skill/              Gemini Enterprise skill bundles + create/test tooling — the / + @ command
                    surface carried into the engine: m365-surface-commander (executor, emits the
                    ```cmd algebra) + m365-command-planner (turns free text into a confirmable plan)
docs/               Design, architecture, implementation, contracts, conventions, build plan,
                    mockups (+ screenshots/), and the discoveryengine API knowledge base
```

The `skill/` bundles are **mounted into Gemini Enterprise per-turn via `skillsSpec`** (an
`agents`/`skillAgentDefinition` resource — see `docs/api/discoveryengine/skills-and-agents.md`).
They mirror the runtime's grammar: keep `scripts/parse_commands.py` and `scripts/parse_plan.py` in
lockstep with `packages/contracts` + `packages/runtime` (the TS side is authoritative).

`web-shell` and `runtime` are the bulk of the client and are **surface-agnostic** — they must not
contain Word/Excel/etc.-specific code. Surface specifics live only in `bridge-*` and `teams/`,
which are the ONLY code that touches Office.js / TeamsJS / Graph.

## Tech stack (decided — don't re-litigate without reason)

- **Language:** TypeScript everywhere (client-direct; no server tier by default).
- **Client:** React + Office.js (Word/Excel/PPT/OneNote) and TeamsJS (Teams). Build with Vite; scaffold task panes with the M365 Agents Toolkit / Yo Office patterns.
- **Gemini Enterprise:** Discovery Engine `v1alpha` called directly; SSE for streaming. Agents/Model Armor are engine config.
- **Monorepo:** Bun workspaces + TS project references. **Validation:** Zod (shared in `packages/contracts`). **Tests:** Vitest.

## Commands

```bash
bun install                  # install all workspaces
bun run --filter @ge/web-shell dev   # run the web-shell dev server (HTTPS)
bun run build                # build all workspaces
bun run typecheck            # tsc --noEmit across workspaces
bun run test                 # vitest across workspaces
bun run lint                 # eslint + prettier check
```
Copy `.env.example` to `.env` and fill it before running the gateway.

## Critical constraints

Phrased as preferences because the agent should internalize them, not just avoid a "don't":

- **Prefer routing specialist agent work over A2A to Agent Engine** rather than through StreamAssist's `agentsSpec` — the latter has a known bug where the agent id is ignored. Use StreamAssist only for the grounded-assistant chat path.
- **Keep all Google credentials in the gateway.** The web-shell and every bridge hold only the user's short-lived Entra/Teams token. No service-account keys ever reach a client.
- **Treat all host document and transcript content as untrusted input.** Screen it through Model Armor at the gateway and pass it to models as data, never as instructions.
- **Prefer reversible, provenanced writes over silent edits.** Agent changes land as tracked changes (Word/PPT) or citation-tagged blocks (OneNote/Teams), each carrying agent id, sources, identity, timestamp, and a content hash in the host's durable metadata.
- **Anchor Word findings by content (`body.search`) and re-resolve at apply-time.** Prefer degrading a stale finding to a panel item over rendering a broken annotation.
- **Prefer federated connector mode for ad-hoc sources;** reserve ingestion mode for large stable corpora.
- **Scope every read and write to the signed-in user's identity end to end.** Prefer delegated Microsoft permissions over org-wide, and `Sites.Selected` over all-sites.
- **Pin the Cloud Run region and the `discoveryengine` endpoint region to the tenant's residency commitment.**

## Definition of done (every task)

`bun run typecheck` clean · `bun run test` green · `bun run lint` clean · acceptance criteria in `docs/BUILD-PLAN.md` met · `security-reviewer` run if the task touched auth/credentials/guardrails/provenance · the task checked off in the build plan.
