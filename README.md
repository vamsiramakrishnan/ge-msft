# Gemini Enterprise for Microsoft 365 — Implementation Package

A Claude Code-ready specification for building a multi-surface Microsoft 365 add-in that brings **Gemini Enterprise** into Word, Excel, PowerPoint, OneNote, and Teams. One stateless gateway, one research unit, one identity, five thin surface clients.

This repo is the **spec and scaffold**, not the implementation — Claude Code builds the implementation from it.

## How to use this with Claude Code

1. Open this folder in Claude Code. It reads `CLAUDE.md` automatically — that's the source of truth.
2. Skim the design references in `docs/` (architecture, design, implementation) and open the UX in `docs/mockups/`.
3. Provision the prerequisites below and copy `.env.example` → `.env`.
4. Start building: run the **`/next-task`** skill. It picks the next unchecked item in `docs/BUILD-PLAN.md`, plans it, implements it against `docs/CONTRACTS.md` and `docs/CONVENTIONS.md`, verifies its acceptance criterion, runs the `security-reviewer` subagent for sensitive code, and checks it off. Repeat.
5. Use `/plan` at the start of each phase and `/verify-surface <name>` before closing a surface's exit gate.

The work is sequenced so the hardest pieces (identity federation, the streaming relay) are proven first in Phase 0's slice-1 spine. If that works, the rest is reuse.

## What's here

```
CLAUDE.md                     The constitution — read first
README.md                     This file
package.json                  npm workspaces monorepo
tsconfig.base.json            strict TS base
.env.example                  config contract
.claude/
  settings.json               permissions + format hook
  skills/next-task/           drive the build plan, one task at a time
  skills/verify-surface/      verify a surface against its mockup + ACs
  agents/security-reviewer.md  reviews identity/credentials/guardrails/provenance
  agents/surface-bridge.md     specialist for the per-surface Office.js/Teams bridges
docs/
  00-surfaces-plan.md         the broader cross-ecosystem strategy (context)
  01-architecture.md          gateway internals, identity federation, anchoring
  02-design.md                the five invariants, per-surface verbs, phasing
  03-implementation.md        packaging, per-surface APIs, repo layout
  CONTRACTS.md                authoritative schemas — implement against this
  CONVENTIONS.md              stack, style, testing, security standards
  BUILD-PLAN.md               the executable checklist Claude Code drives
  mockups/                    the clickable UX spec, one per surface
manifests/
  m365-unified.manifest.json  Package A: Word/Excel/PPT task panes + Teams (tab/bot/meeting/ME)
  onenote.manifest.xml        Package B: OneNote (web-only, legacy manifest)
packages/  services/          where Claude Code builds (see CLAUDE.md for the map)
```

## Prerequisites to provision

- **Node 20+** and npm.
- A **Google Cloud project** with Gemini Enterprise (an app/engine), Agent Engine, Discovery Engine, Model Armor, and a **Workforce Identity Pool** federating your Entra tenant.
- A **Microsoft Entra tenant** with: a client (NAA) app registration for the add-in, a gateway app registration, and a connector app registration for SharePoint/OneDrive (prefer delegated permissions and `Sites.Selected`).
- A **NotebookLM Enterprise** notebook (the curated core of the research unit) — included with Gemini Enterprise Standard/Plus.
- For Teams: an **Azure Bot** registration for the bot/message-extension.

## The one-line shape

Build the gateway, the unit resolver, the identity federation, and the web-shell once; express them as Word/Excel/PowerPoint task panes and a Teams tab/meeting/bot in one package (plus a companion OneNote add-in), where each surface adds only a thin content bridge — and the same backend later carries the whole thing into Salesforce and SAP.
