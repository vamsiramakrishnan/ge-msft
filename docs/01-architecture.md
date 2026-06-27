# Architecture — Gemini Enterprise in Word

> **⚠️ Superseded in part — updated by `ADR-0001` (client-direct) and `ADR-0003`–`ADR-0006`.**
> This doc describes the **original gateway-centric** three-tier design: a stateless Cloud Run
> **Surface Gateway** holding Google credentials, federating identity server-side, screening with
> Model Armor, routing, relaying, and signing provenance. **That tier no longer exists.** The
> project is now **client-direct**: identity federates to Google in the browser (WIF), the add-in
> calls Discovery Engine directly, and Model Armor / agent routing / grounding are **engine config**,
> not our code. Read this doc for the *anchoring contract* (§7, still authoritative — content anchors
> + apply-time re-resolution) and the *identity-federation reasoning* (§5, now done client-side); for
> everything else the **ADRs are the source of truth**. The assist read/write loop is no longer an
> intent dispatch — it is the **CLI command protocol → composable algebra** (`ADR-0004`/`ADR-0005`).

**A buildable design for the Word add-in, and the gateway underneath it that every other surface reuses.**
*Companion to the Surfaces plan and the click-through prototype. Architecture-of-record: the ADRs.*

---

## 1. The shape

The mistake would be to architect "a Word add-in." What you're actually building is a **Surface Gateway** — one stateless service that turns Gemini Enterprise into something any host can consume — and then a *thin Word client* on top of it. The Word add-in is the first adapter; Excel, Salesforce, and SAP are later adapters on the same core. This is the build-once thesis from the plan, made concrete: the expensive parts (identity federation, guardrails, streaming relay, provenance, audit) live in the gateway and are written exactly once.

Three tiers, three planes:

```
        CLIENT TIER                 GATEWAY TIER                    INTELLIGENCE TIER
 ┌────────────────────────┐   ┌──────────────────────────┐   ┌───────────────────────────┐
 │  Word task pane        │   │   SURFACE GATEWAY        │   │   Gemini Enterprise       │
 │  (Office.js + React)   │   │   (Cloud Run, stateless) │   │                           │
 │                        │   │                          │   │  StreamAssist (assistant) │
 │  • DocBridge (R/W)     │   │  • Token validation      │   │  Specialist agents (A2A)  │
 │  • AnnotationManager   │◄─►│  • Identity federation   │◄─►│   on Agent Engine         │
 │  • AuthClient (NAA)    │SSE│  • Router (assist|agent) │   │   (Review, Redline, …)    │
 │  • StreamClient        │   │  • Streaming relay       │   │  Data stores / connectors │
 │  • ProvenanceStore     │   │  • Model Armor + audit   │   │  Model Armor · Memory Bank│
 └────────────────────────┘   │  • Provenance signing    │   └───────────────────────────┘
   runs inside Word's          └──────────────────────────┘
   webview · no GCP creds        the only thing that holds
                                 GCP credentials
 ───────────────────────────────────────────────────────────────────────────────────────
   CONTROL PLANE  manifest · distribution (AppSource / M365 admin) · agent registration
   DATA PLANE     request/response · streaming · doc context in, tokens + findings out
   TRUST PLANE    NAA → Workforce Identity Federation · Model Armor · provenance · residency
```

The single most important boundary: **the client never holds Google *secrets*.** (Still true. What
changed: identity is now exchanged **in the browser** via WIF, not server-side — the client holds the
user's short-lived Entra token and the federated Google access token derived from it, in memory only.
The "gateway tier" in the diagram above is gone; see `ADR-0001`.) Everything else follows from that.

---

## 2. The client (task pane)

Office.js + React in the Word webview, five modules with clean seams so the same modules port to Excel/PowerPoint later:

- **DocBridge** — all Office.js reads and writes. Reads: selection, body text, structured ranges, content controls, and `getFileAsync` for the *rendered* document when multimodal grounding is needed. Writes: tracked changes, comment replies, content-control population, annotations. This is the only module that touches the document model.
- **AnnotationManager** — turns agent findings into Word annotations (the annotations API) and back. Owns the anchor-resolution problem (§7) and the accept/reject lifecycle.
- **AuthClient** — MSAL.js with Nested App Authentication. `acquireTokenSilent` for the signed-in user's Entra token, brokered by the Office host; falls back to the Office Dialog API when SSO is unavailable. Never sees Google.
- **StreamClient** — consumes SSE from the gateway, with a chunked-polling fallback abstracted behind one interface so callers don't care which transport is live.
- **ProvenanceStore** — reads/writes the custom XML part embedded in the .docx (agent id, sources, hash, identity, timestamp, and the StreamAssist session id so a reopened document resumes its conversation).

The client is deliberately dumb about intelligence. It renders, it reads/writes the document, it authenticates the user. All reasoning, routing, and trust enforcement happen one tier down.

---

## 3. The Surface Gateway (the heart)

A stateless Cloud Run service. It exists for three reasons that can't be satisfied client-side: you can't safely hold GCP credentials in a webview, identity federation must happen server-side, and you want **one** enforcement point for guardrails, residency, audit, and routing across every surface.

Responsibilities, in request order:

1. **Authenticate the caller.** Validate the inbound Entra JWT against Entra's JWKS (issuer, audience, expiry). Extract identity and claims — `groups`/`roles` matter, because they become the agent's data scope.
2. **Federate identity.** Exchange the Entra OIDC token for Google credentials via Workforce Identity Federation (Security Token Service), against a workforce pool that trusts the M365 tenant. This is §5 — the hard part.
3. **Screen the input.** Run Model Armor over the document context flowing in. The document is *untrusted* — it can contain injected instructions — so it's treated as data, never as prompt, and screened before it reaches a model.
4. **Route.** Decide between StreamAssist (grounded assistant chat) and a specialist A2A agent on Agent Engine (review, redline, compliance). The router is the surface-agnostic brain; the Word client just names an intent (`assist`, `review`, `resolve-comment`, `regen-clause`).
5. **Relay the stream.** Consume the server-stream from StreamAssist or the A2A response and re-emit as SSE, appending grounding citations as they arrive. Backpressure-aware; refreshes the federated token if it expires mid-stream.
6. **Sign provenance.** Attach grounding sources, the agent id, identity, timestamp, and a content hash to each response so the client can persist them into the document's custom XML.
7. **Audit.** Structured log of every invocation — who, which agent, which sources, what changed — to Cloud Logging and BigQuery for compliance.

Because routing and federation are surface-agnostic, the Salesforce and SAP adapters from the plan are *new entry handlers on this same service*, not new systems. That's the leverage.

---

## 4. The Gemini Enterprise side

- **StreamAssist** — the grounded assistant for panel chat, grounding simultaneously on the live document (passed as file-context) and the connected enterprise data stores. The streaming, session-aware entry point.
- **Specialist agents on Agent Engine** — ADK agents exposed as A2A services (Contract Review, Redline, Compliance). Built with ADK (adk-fluent fits cleanly here as the authoring layer). No-code Agent Designer agents reach the gateway through the thin A2A shim from the plan. Routing to these over A2A — rather than through StreamAssist's `agentsSpec` — also sidesteps the early-2026 `agentsSpec` bug entirely.
- **Data stores / connectors** — the enterprise grounding: vendor risk policies, prior contracts, regulatory texts (CPS 234 and the like), via Discovery Engine data stores.
- **Model Armor** for safety, **Memory Bank** for cross-session agent memory where an agent needs to remember a deal across documents.

One judgment call worth flagging: a full multi-clause review pass is a multi-agent, possibly minutes-long job. That's a candidate for durable execution (your Tape work) rather than a single synchronous request — kick off the review, stream findings back as each clause completes, and survive a dropped connection.

---

## 5. Identity federation — the part that actually decides whether this ships

Everything technical here is the same problem three times across surfaces, and it's where an enterprise buyer pushes hardest. The flow for Word:

```
 Word add-in            Office host         Gateway              Google STS         Gemini Enterprise
     │                      │                  │                    │                     │
     │ acquireTokenSilent   │                  │                    │                     │
     ├─────────────────────►│ broker token     │                    │                     │
     │◄─────────────────────┤ (Entra OIDC)     │                    │                     │
     │  POST /review  + Bearer(Entra token)    │                    │                     │
     ├────────────────────────────────────────►│                    │                     │
     │                      │     validate JWT (Entra JWKS)         │                     │
     │                      │                  │ STS token exchange │                     │
     │                      │                  ├───────────────────►│                     │
     │                      │                  │  federated Google  │                     │
     │                      │                  │◄───────────────────┤  token (scoped)     │
     │                      │                  │   call agent / StreamAssist (as identity)│
     │                      │                  ├─────────────────────────────────────────►│
     │◄═══════════ SSE: findings + citations ══┤◄═══ stream + grounding ══════════════════┤
```

The design decision underneath it: **does the agent act as the end user, or as a service principal with the user's claims?** Two options, and you'll likely use both:

- *End-user-scoped* — federate to a principal whose data-store ACLs mirror the user's, so document-level access controls hold. Highest fidelity, needed for regulated data.
- *Service principal + attribute-based access* — run as a service account but derive data filters from the user's `groups`/`roles` claims passed through. Simpler, fine for less sensitive grounding.

Get this wrong in either direction and you either over-expose data (an over-privileged service account answering for everyone) or break the experience (no access at all). The federated token is cached with a TTL safely under its expiry; on a `401` the client silently refreshes the NAA token and retries the read.

---

## 6. The runtime flows

**Grounded chat (intent: `assist`).** DocBridge grabs the selection (and body if needed) → client POSTs `{sessionId, query, docContext}` → gateway validates, federates, screens, calls StreamAssist with the doc as file-context → tokens + grounding stream back as SSE → panel renders streaming text and citation chips → ProvenanceStore writes the session id into the .docx.

**Inline review (intent: `review`).** Client sends the whole document (text, or the rendered file via `getFileAsync` for multimodal) → gateway routes to the Review agent over A2A → agent grounds on the doc + policy stores and returns a structured findings list → gateway screens and hashes → AnnotationManager resolves each finding to a Word range and renders squiggles + hover cards → on Accept, DocBridge applies a tracked change and ProvenanceStore records that finding's provenance.

**Comment resolve (intent: `resolve-comment`).** A comment event (or user action) sends `{commentText, anchoredRange, docContext}` → agent returns `{editedText, replyText, sources}` → client applies the edit as a tracked change, posts `replyText` as a threaded reply, resolves the comment, writes provenance.

**Surgical regen (intent: `regen-clause`).** The clause is an addressable content control; the request is scoped to that control's id → the agent rewrites only that text → it streams into the control while the client visibly locks the others. The content control boundary is what makes "regenerate just this" a clean operation instead of a diff-and-pray.

---

## 7. The one non-obvious decision: how findings anchor to the document

This is the contract that separates a real build from the prototype. The agent returns findings, but **the agent cannot know Word's internal range identity** — paragraph `uniqueLocalIds` are GUIDs that differ across sessions and coauthors. So the agent↔client contract anchors on *content*, and the client resolves to a range at apply-time:

```jsonc
// Agent → gateway → client
{
  "findings": [{
    "category": "policy",                    // style | policy | ground
    "matchText": "99.5% of the time",        // the exact text to locate
    "contextHint": "Services are available", // disambiguates repeats
    "title": "Availability below FSI standard",
    "why": "...",
    "suggestion": "99.9% of the time",
    "sources": [{ "title": "Vendor Risk Policy v4 §3.2", "uri": "..." }],
    "confidence": 0.91,
    "hash": "a1f9c4e2"                        // for provenance
  }]
}
```

The client uses `body.search(matchText)` to locate the range, disambiguates with `contextHint` if there are multiple hits, and *re-resolves at the moment of Accept* — because the user may have edited the document between the request and the response. If the matched text no longer exists (the user already changed it), the finding **degrades gracefully into a panel item** instead of a broken inline annotation. Designing for anchor drift up front is the difference between a demo and a tool people leave running while they type. The same rule governs coauthoring: edits and annotation events carry `source: Local | Remote`, so the client ignores the agent's own writes and never enters a feedback loop.

---

## 8. UX feature → service mapping

| UX capability | Client mechanism | Gateway / Gemini |
|---|---|---|
| Inline annotations | Annotations API + `body.search()` anchoring | Review agent (A2A) returns findings |
| Comment task queue | Comment events + tracked changes + replies | `resolve-comment` → agent edit + reply |
| Surgical regeneration | Content controls (addressable, lockable) | `regen-clause` scoped to one control |
| Provenance | Custom XML part (read/write) | Gateway signs sources + agent id + hash |
| Ambient grounding | Selection/paragraph events, debounced | `assist` with current section as context |
| Multimodal grounding | `getFileAsync` (PDF/OOXML) | Multimodal Gemini sees rendered layout |
| Voice | Live API over WebSocket | Ephemeral token minted by gateway |
| Excel cells (later) | Streaming custom functions / linked-entity load service | Same `assist` / retrieval endpoints |

Voice is the one path that doesn't relay through the gateway for the audio stream — the Live API wants a near-direct WebSocket for latency, so the gateway mints a short-lived scoped token and the client connects to Live directly. Everything else flows through the gateway.

---

## 9. Trust plane

- **Untrusted document.** The document body is hostile input by default. Model Armor screens it on the way in; the gateway passes it as data, never as instruction. This is the prompt-injection defense the M365 docs themselves warn about.
- **Provenance that travels.** Every generated or edited paragraph carries agent, sources, identity, timestamp, and content hash in the custom XML part — auditable by any reviewer later, with no external lookup. (Your ISO 26262 instinct, applied to documents: traceability as a property of the artifact.)
- **Residency.** Pin the Cloud Run region and the `discoveryengine` endpoint region to the customer's commitment, and map the M365 tenant region to the matching Google region. For JAPAC, Australian customer data stays in-region end to end; document content never crosses a region boundary.
- **Audit.** Structured invocation log → BigQuery: who asked, which agent answered, which sources grounded it, what changed in the document.

---

## 10. Failure modes and contracts

- **SSE unavailable or dropped** → StreamClient falls back to chunked polling behind the same interface.
- **`agentsSpec` bug** → specialist work routes to Agent Engine over A2A, not through StreamAssist's agent spec; the architecture avoids the bug rather than waiting on the fix.
- **Anchor drift** → re-resolve at Accept; degrade to a panel item if the text is gone (§7).
- **Token expiry mid-stream** → gateway refreshes the federated token; client refreshes NAA on `401` and retries idempotent reads.
- **Write idempotency** → every apply carries a change id so a retry can't double-insert.
- **Large documents** → never inline OOXML; use StreamAssist file-context or transient retrieval, with a hard context cap.
- **Coauthoring loops** → ignore `source: Remote`/own edits as above.

---

## 11. Build order (thin vertical slices)

1. **Spine.** Task pane + NAA auth + gateway + StreamAssist grounded chat over the selection, streaming. Proves the whole chain end to end: user identity → federation → Gemini → stream back. Nothing fancy, but if this works, the hard parts are solved.
2. **Signature.** Review agent (A2A) + annotations API + anchor resolution + accept-as-tracked-change. The inline experience that differentiates the product.
3. **Trust.** Provenance (custom XML) + the comment-resolve loop.
4. **Depth.** Surgical regeneration (content controls) + multimodal + voice.
5. **Hardening.** Residency pinning, Model Armor, audit to BigQuery, distribution via M365 admin Integrated Apps, admin governance.

Slice 1 is the real risk-retirement: it's the smallest thing that exercises identity federation and the streaming relay, which are the two pieces most likely to be hard.

---

## 12. One sentence

Build a stateless Surface Gateway that federates the user's Microsoft identity to Google, screens the document as untrusted input, and relays Gemini's streamed reasoning and grounded findings back to a thin Office.js client that anchors them to the live document by content rather than by ID — and the same gateway is what later carries Gemini into Excel, Salesforce, and SAP.
