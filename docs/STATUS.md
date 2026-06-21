# Status — what's built (and what's next)

The honest, current state of the codebase. Companion to `BUILD-PLAN.md` (the original checklist)
and `CAPABILITY-MAP.md` (the I/O inventory). Updated as of the client-direct reorg.

> **Architecture:** client-direct (see `ADR-0001`). The add-in federates the signed-in user's
> Entra identity to Google (Workforce Identity Federation, in the browser) and calls Gemini
> Enterprise (Discovery Engine `v1alpha`) directly. **No gateway by default**; the only optional
> server piece is a transparent CORS/audit proxy via `proxyUrl`. Model Armor, agent routing, and
> grounding are Gemini Enterprise engine config, not our code.

## Verification baseline

`npm run typecheck` clean · **216 tests across 32 files green** (Vitest) · `npm run lint` clean.

## Packages — built vs planned

| Package | Role | Status |
|---|---|---|
| `@ge/contracts` | Shared types + Zod schemas (the core↔bridge boundary) | ✅ built |
| `@ge/content` | Native-first content processing: object model → blocks → budgeted chunks | ✅ built |
| `@ge/gemini-client` | Client-direct Discovery Engine: WIF exchange, `streamAssist`, `search`, `completeQuery`, `checkGrounding`, `rank` | ✅ built |
| `@ge/graph-client` | Microsoft Graph reader (Plane B / estate): messages, events, driveItems, users, `/search` | ✅ built |
| `@ge/triggers` | Event-driven layer: `HostEvent` lifecycle, `TriggerRegistry`, debounce, the actuation gate | ✅ built |
| `@ge/runtime` | Surface-agnostic core: `DocBridge`/`AuthClient`, `AssistSession`, `ContextModel`, `Orchestrator` | ✅ built |
| `@ge/web-shell` | App core: `NaaAuthClient` (NAA), `composeSession`, `PanelController`, `ProvenanceStore`, host detection | ✅ core built (React/Vite/manifest shell pending) |
| `@ge/bridge-word` | Word: native capture + content-anchored tracked changes + `watch()` | ✅ built |
| `@ge/bridge-excel` | Excel: range capture + address-anchored `write-cells` + `watch()` | ✅ built |
| `@ge/bridge-outlook` | Outlook: mail capture + reviewable reply + the **on-send gate** | ✅ built |
| `@ge/teams` | Teams: transcript capture + reviewable post-message + meeting events | ✅ built |
| `@ge/bridge-powerpoint` | PowerPoint: deck composer + speaker notes | ⬜ planned stub |
| `@ge/bridge-onenote` | OneNote page synthesis (web-only, legacy manifest) | ⬜ planned stub |

The `services/gateway` tier was **removed** in the client-direct reorg — it contradicted ADR-0001.

## Capabilities delivered

- **Identity, client-direct.** `NaaAuthClient` (MSAL Nested App Auth) yields the Entra id token (WIF
  subject), delegated Graph tokens (Plane B), and the identity for provenance. `WifTokenClient`
  exchanges Entra→Google (RFC 8693) in the browser, cached with TTL/skew, single-flight, epoch-safe
  invalidation. No Google credential ever reaches a client.
- **Grounded assist loop.** `AssistSession` ties a bridge to `streamAssist`: attach context →
  stream a grounded answer (tokens + citations + provenance) → apply a reversible, provenanced
  actuation. Session id is captured and resumable across surfaces.
- **Foundational retrieval (beyond assist).** `search` (faceted/filtered, boost, snippets,
  pagination, `dataStoreSpecs`), `completeQuery` (type-ahead), `checkGrounding` (per-claim score —
  backs the on-send / pre-actuation gate, fail-closed), `rank` (semantic rerank). Search hits map to
  reference context (reference-over-inline).
- **Native-first content processing.** Office object model → typed `Block`s with host locators →
  token-budgeted, section-aware chunks → contextualized → `ResolvedContext` mapping 1:1 to
  `query.parts`. Budget picks inline / reference / upload-for-code-execution.
- **Two context planes.** Plane A (in-document, Office.js, per bridge) and Plane B (estate, Graph:
  search SharePoint/OneDrive/mail/calendar/people as the user).
- **Event-driven, not assistant-spamming.** `watch()` on each bridge emits `HostEvent`s (selection /
  document / comment changes with coauthor origin; Outlook item-change + on-send; Teams meeting-end).
  The `Orchestrator` debounces and routes. Most events **construct context** via the `ContextModel`
  (cheap, no model call); the `TriggerRegistry` gate handles the rare protective moments
  (on-send grounding veto, pre-actuation veto); suggestions are scarce and ignorable.
- **Context construction → session commit.** Events build a compact, capped, data-framed working
  brief that is **folded** into the next turn (no extra call) or **primed** at a checkpoint
  (meeting-ended). Once sent it is resident in the Gemini Enterprise session and not re-sent;
  version-scoped so nothing is lost or double-sent across failures.
- **Reversible, provenanced writes.** Every actuation carries agent id, sources, identity, timestamp,
  and a content hash; Word anchors by content (`body.search`, re-resolved at apply-time, degrades on
  drift); Outlook/Teams open reviewable forms rather than sending silently. `PanelController` stages
  proposals for explicit confirmation and records outcomes in `ProvenanceStore`.

## What's next (pending)

1. **Sideload shell** — the React task pane + Vite build + unified manifest over the `web-shell`
   core (the last mile to actually load in a host).
2. **PowerPoint + OneNote bridges** — real surfaces in the vision; currently planned stubs.
3. **Graph change-notification source** — `estate-changed` events (subscriptions/delta).
4. **Security hardenings noted by review** — keep `decideSend` total; add a bounded timeout so a
   *hung* on-send trigger cannot wedge Send; validate `location`/`proxyUrl` at config construction.
5. **A2UI renderer** and the **v1 `addContextFile`** path (code-execution uploads) — designed, not built.

See `MICROSOFT-ADDIN-CAPABILITIES.md` for the Microsoft 365 add-in surface we build on, and
`CAPABILITY-MAP.md` for the per-capability read/write inventory.
