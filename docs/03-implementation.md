# Implementation — Gemini Enterprise M365 Add-in Suite

**The buildable spec: how the one-gateway/one-unit design becomes five shipping surfaces.**
*Companion to the architecture doc (gateway internals, identity federation, anchoring) and the design doc (experience, phasing).*

---

## 1. Architecture recap

One stateless **Surface Gateway** on Cloud Run holds all Google credentials, performs identity federation, screens input, routes between StreamAssist and A2A specialist agents on Agent Engine, relays streams as SSE, and signs provenance. Every client is thin: it renders, reads/writes the host's content, authenticates the user, and names an intent (`assist`, `review`, `resolve-comment`, `regen-clause`, `draft-slides`, `synthesize`, `meeting-notes`). The full internals are in the architecture doc; this document covers the *client* tier across five surfaces and the packaging that unifies them.

The unifying engineering fact: **the same web app is the client for most surfaces.** Microsoft's packaging lets one distributable bundle carry Word/Excel/PowerPoint/Outlook task panes *and* Teams tab/bot/meeting capabilities, and lets a single page experience appear as both an add-in task pane and a Teams personal tab. So the panel, the unit composer, the auth client, and the stream client are written once and reused; only the per-surface document bridge differs.

---

## 2. Packaging and surface targeting

Two packages, because OneNote is the exception:

- **Package A — the unified M365 bundle** (TeamsJS 2.19+, unified app manifest 1.13+, built with the M365 Agents Toolkit). Declares: Word + Excel + PowerPoint task panes (and Outlook if desired) *and* Teams capabilities (personal tab, meeting app, bot, message extension). One app, one install, multiple surfaces. This is the spine of the suite.
- **Package B — the OneNote add-in** (legacy add-in-only XML manifest, Yo Office). OneNote runs only in OneNote on the web and does not use the unified manifest, so it ships as a companion add-in that reuses the same web app and gateway.

```
 ┌───────────────── Package A · unified M365 manifest ─────────────────┐   ┌── Package B ──┐
 │  Word TP   Excel TP   PowerPoint TP   │  Teams tab · meeting · bot  │   │  OneNote add-in│
 │  (Office.js)          (Office.js)     │  (TeamsJS + Bot Framework)  │   │  (OneNote API) │
 └───────────────────────┬───────────────────────────┬────────────────┘   └───────┬───────┘
                          └──────────  SAME WEB APP (panel · unit composer · auth · stream)  ──┘
                                              │
                                     Surface Gateway (Cloud Run)
```

Manifest essentials: `validDomains` includes the gateway/host domain (or task panes open in a separate window on desktop); SSO redirect URIs for NAA (the brokered `brk-multihub://` plus SPA fallback); Teams capabilities declare the bot ID and message-extension commands; resource-specific consent (RSC) for the meeting app's transcript access.

---

## 3. Per-surface client implementation

Shared across all: **DocBridge** (host content R/W), **AuthClient** (NAA/TeamsJS SSO), **StreamClient** (SSE + polling fallback), **ProvenanceStore**, **UnitComposer** (the notebook + connectors panel control). What changes per surface is the DocBridge implementation and the signature feature.

**Word** — `Word.run`. Read: selection, body, content controls, `getFileAsync` (rendered doc for multimodal). Write: tracked changes (`changeTrackingMode`), `insertOoxml`, comment replies, content-control population, **annotations** (`onAnnotation*` events + the annotations API) for inline squiggles with hover cards and popup actions. Findings anchor by content (`body.search(matchText)`) and re-resolve at apply-time — see the architecture doc's anchoring contract. Provenance → custom XML part.

**Excel** — two mechanisms beyond the panel. **Streaming custom functions** (`@streaming`, `setResult` called repeatedly) implement `=GE.ASK(prompt, range)` — the function calls the gateway and streams the grounded answer into the cell. **Linked-entity cells** (`LinkedEntityCellValue` + a linked-entity load service) implement vendor/entity cells: the load service calls the gateway/Gemini retrieval to return up-to-date entity values on demand, expandable as cards, without bloating the workbook. Core grid R/W via `Excel.run`; `onChanged` range events for recompute.

**PowerPoint** — `PowerPoint.run`. Read slide text and selected shapes for grounding; write generated slides (`insertSlidesFromBase64` / slide and shape insertion), speaker notes, and layout changes. The deck composer streams slides into the deck as the agent returns them; each slide records its source provenance.

**OneNote** — `OneNote.run` via the `Application` object (Notebook / Section / Page / Outline / Paragraph). Read the page's sources; write synthesised content as page outlines with inline citation tags. The notebook *is* the unit here, so the NotebookLM overview actions (audio/video) call the gateway's NotebookLM endpoints scoped to this notebook. Web-only; narrower API — scope accordingly.

**Teams** — three capabilities, all reusing the same web app where possible:
- *Meeting app* (in-meeting side panel) — the same panel, hosted as a meeting-stage/side-panel tab via TeamsJS, grounding on the unit **plus the live transcript** (RSC consent) to produce live notes and grounded action items.
- *Bot* (Bot Framework / Azure Bot) — conversational "ask the agent" in meeting chat or channels; proxies to the gateway; renders responses as Adaptive Cards (recap card, action-item card).
- *Message extension* — "ground this message on the unit" from the compose box or as a search command.

---

## 4. Identity and auth

One model, two front doors:

- **Office surfaces (Word/Excel/PowerPoint/OneNote)** — MSAL.js with **Nested App Authentication**: `acquireTokenSilent` for the signed-in user's Entra token, brokered by the Office host, with an Office Dialog API fallback when SSO is unavailable.
- **Teams surfaces** — **TeamsJS SSO** (`getAuthToken` / MSAL) for the tab and meeting app; the bot authenticates via its Entra app registration and the Bot Framework token service.

Both front doors produce an Entra token that hits the gateway, which validates it and exchanges it through **Workforce Identity Federation** for scoped Google credentials, so Gemini — and the federated SharePoint/OneDrive connectors behind the unit — act with the user's identity. This closes the identity envelope: the *same* user identity scopes the SSO into the app, the grounding in Gemini, and the federated connector reads/writes. One registration strategy, least privilege throughout (delegated over org-wide; `Sites.Selected` over all-sites). The federated token is cached under its TTL; a `401` triggers a silent NAA/TeamsJS refresh and an idempotent retry.

---

## 5. The research unit, in code

The unit is resolved server-side per request so clients stay dumb. The client sends a unit descriptor; the gateway resolves it:

```jsonc
// client → gateway, on every grounded call
{
  "intent": "review",
  "unit": {
    "notebookId": "nb_vendor_risk_7f3",        // the curated core (NotebookLM)
    "connectors": [                              // the live, federated edge
      { "type": "sharepoint", "mode": "federated", "scope": "sites/DealRoom" },
      { "type": "onedrive",   "mode": "federated" }
    ],
    "surfaceContext": { "kind": "word", "selection": "..." }  // the working surface
  }
}
```

Resolution rules the gateway applies: **federated** connectors fetch live at query time using the user's identity (nothing copied; results reflect only what the user can see); **ingestion**-mode sources are queried from their pre-built index for speed on large stable corpora. The **notebook** is the precision instrument — for regulated work the gateway can restrict grounding to the notebook ("answer only from these sources"). Connector **actions** (write-back: upload the redlined doc, check in/out) are separate, explicitly-authorised calls, never implicit in a grounding request. One data type per data store, and only one actions-enabled data store per connector type — provision accordingly.

---

## 6. Streaming, provenance, guardrails (applied per surface)

These are gateway responsibilities (detailed in the architecture doc); each surface consumes them through the shared StreamClient and ProvenanceStore:

- **Streaming** — SSE from the gateway into whatever the surface renders: panel text (all), cell value (Excel streaming function), slide rail (PowerPoint), page block (OneNote), meeting notes (Teams). Polling fallback behind one interface.
- **Provenance** — agent id, sources, identity, timestamp, content hash attached to every result; persisted into the host's durable metadata (custom XML in Word/PPT; cell notes/linked-entity metadata in Excel; page metadata in OneNote; the recap card and channel post in Teams).
- **Guardrails** — Model Armor screens host content (untrusted) on the way in; outputs and write-back actions are screened; everything is audited to BigQuery. The document/transcript is data, never instruction.

---

## 7. Repo structure

```
/packages
  /web-shell        ← the panel, UnitComposer, AuthClient, StreamClient, ProvenanceStore (shared)
  /bridge-word      ← Word DocBridge + annotations/anchoring
  /bridge-excel     ← custom functions + linked-entity load service
  /bridge-powerpoint← deck composer + notes
  /bridge-onenote   ← OneNote page synthesis (separate package/manifest)
  /teams            ← tab/meeting host + bot (Bot Framework) + message extension
  /manifests        ← unified M365 manifest (A) + OneNote XML manifest (B)
/services
  /surface-gateway  ← Cloud Run: auth, federation, router, relay, provenance, audit
  /agents           ← ADK A2A specialist agents (review, redline, compliance) for Agent Engine
/infra              ← Workforce Identity Pool, Secret Manager, Cloud Run, Azure Bot, residency config
```

The discipline that makes this tractable: `web-shell` is the bulk of the client and is written once; each `bridge-*` is small and surface-specific; `surface-gateway` and `agents` are the same regardless of surface (and the same that will later serve Salesforce/SAP).

---

## 8. Build sequence (aligned to the design phases)

1. **P0 Foundation** — `surface-gateway` (auth → federation → StreamAssist → SSE), the unit resolver, `agents` (one A2A review agent), provenance signing, and `web-shell`. Prove slice 1: grounded streamed answer over a Word selection. This retires identity federation and the streaming relay — the two hard pieces.
2. **P1 Word** — `bridge-word`: annotations, anchoring, comment queue, surgical regen, provenance into custom XML. Ship to a design partner.
3. **P2 Excel + PowerPoint** — `bridge-excel` (custom functions + load service) and `bridge-powerpoint` (composer + notes). Gateway/unit/identity unchanged.
4. **P3 OneNote + Teams** — `bridge-onenote` (Package B) and `/teams` (tab + meeting app + bot + message extension; Azure Bot infra). Most divergent client work.
5. **P4 Continuity + scale** — cross-surface provenance/unit continuity, residency pinning, observability/eval, AppSource + Teams Store + admin-center distribution, then replicate.

---

## 9. Testing, distribution, governance

- **Testing** — Office surfaces sideload via the Agents Toolkit / Yo Office; Teams runs in the web client via the toolkit; cross-platform validation is required for marketplace (an add-in must work everywhere it declares support). Script Lab for rapid Office.js prototyping of each bridge.
- **Distribution** — Package A via AppSource (public) or, more likely for enterprise customers, centralized deployment through Integrated Apps in the M365 admin center; Teams capabilities surface through the same package. Package B (OneNote) via AppSource or admin deployment. IT admins gate availability and consent via the Package Management API and Teams admin controls.
- **Governance** — least-privilege Entra app registrations; admin consent flows; RSC for Teams meeting transcript; per-tenant residency pinning (Cloud Run region + `discoveryengine` endpoint region matched to the tenant); full invocation audit to BigQuery.

---

## 10. Risk and dependency register

- **`streamAssist` `agentsSpec`** — route specialist work to Agent Engine over A2A, not through assistant agent-spec, sidestepping the early-2026 bug.
- **OneNote constraints** — web-only, legacy manifest, thinner API; scope its experience to the API and ship it as a separate package.
- **Teams bot infrastructure** — Azure Bot + Bot Framework + RSC consent is real additional surface; budget it in P3.
- **Anchor drift (Word)** — re-resolve findings at apply-time; degrade to panel items when matched text is gone.
- **Connector ACL intersection** — the unit must enforce the intersection of notebook and connector permissions for the signed-in user; the identity envelope is the control.
- **Marketplace lead time** — start AppSource/Teams Store certification early.
- **Residency on the broker hop** — federated connectors call Microsoft endpoints outside Google's network (VPC SC doesn't wrap them); mitigate with VPC firewall to specific FQDNs and keep the notebook's data in-project/in-region.
- **SDK currency** — keep agents on the current Gemini Enterprise Agent Platform SDK; TeamsJS 2.19+ and unified manifest 1.13+ for cross-host.

---

## 11. One sentence

Write the gateway, the unit resolver, the identity federation, and the web-shell once; express them as Word/Excel/PowerPoint task panes and a Teams tab/meeting/bot in a single unified package (plus a companion OneNote add-in), where each surface adds only a thin content bridge — and the same backend later carries the whole thing into Salesforce and SAP.
