# CLAUDE.md — Gemini Enterprise for Microsoft 365

This file is the source of truth for how this repo works. Read it fully before doing anything.

## What we're building

A multi-surface Microsoft 365 add-in that brings **Gemini Enterprise** into Word, Excel, PowerPoint, OneNote, and Teams. One stateless **Surface Gateway** holds all Google credentials and does the real work; the surface clients are thin. The agent grounds on a composable **research unit** (a NotebookLM notebook + federated SharePoint/OneDrive sources + the working document), acts with the **signed-in user's identity end to end**, and leaves changes that are **traceable and reversible**.

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

```
packages/
  contracts/        Shared TypeScript types + Zod schemas (the gateway↔client boundary)
  web-shell/        The reused web app: panel, UnitComposer, AuthClient (NAA), StreamClient, ProvenanceStore
  bridge-word/      Word DocBridge: annotations, anchoring, comment queue, surgical regen
  bridge-excel/     Excel: =GE.ASK streaming custom function + linked-entity load service
  bridge-powerpoint/PowerPoint: deck composer + speaker notes
  bridge-onenote/   OneNote page synthesis (ships as its own package — web-only, legacy manifest)
  teams/            Teams tab + meeting app + bot (Bot Framework) + message extension
services/
  gateway/          Cloud Run: auth, identity federation, router, SSE relay, provenance, audit
  agents/           ADK (Python) A2A specialist agents for Agent Engine (review, redline, compliance)
manifests/          m365-unified.manifest.json (Package A) + onenote.manifest.xml (Package B)
docs/               Design, architecture, implementation, contracts, conventions, build plan, mockups
```

`web-shell` is the bulk of the client and is **surface-agnostic** — it must not contain Word/Excel/etc.-specific code. Surface specifics live only in `bridge-*` and `teams/`.

## Tech stack (decided — don't re-litigate without reason)

- **Language:** TypeScript everywhere on the client/gateway; Python for `services/agents` (ADK).
- **Client:** React + Office.js (Word/Excel/PPT/OneNote) and TeamsJS + Bot Framework (Teams). Build with Vite; scaffold task panes with the M365 Agents Toolkit / Yo Office patterns.
- **Gateway:** Fastify on Node 20+, deployed to Cloud Run. SSE for streaming.
- **Agents:** Google ADK, exposed as A2A servers, deployed to Agent Engine.
- **Monorepo:** npm workspaces. **Validation:** Zod (shared in `packages/contracts`). **Tests:** Vitest (TS), pytest (agents).

## Commands

```bash
npm install                  # install all workspaces
npm run dev -w services/gateway     # run the gateway locally (needs .env)
npm run dev -w packages/web-shell   # run the web-shell dev server (HTTPS)
npm run build                # build all workspaces
npm run typecheck            # tsc --noEmit across workspaces
npm run test                 # vitest across workspaces
npm run lint                 # eslint + prettier check
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

`npm run typecheck` clean · `npm run test` green · `npm run lint` clean · acceptance criteria in `docs/BUILD-PLAN.md` met · `security-reviewer` run if the task touched auth/credentials/guardrails/provenance · the task checked off in the build plan.
