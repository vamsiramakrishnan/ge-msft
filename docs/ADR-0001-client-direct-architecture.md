# ADR-0001 — Client-direct add-in (no custom gateway by default)

**Status:** Accepted (2026-06-21) · supersedes the gateway-centric framing in `01-architecture.md`/`BUILD-PLAN.md` Phase 0.

The API observations below record the June decision. Current public/widget skill routing and
request construction live in `packages/gemini-client`; later route-specific skill support does not
change this client-direct decision. See [the current contract map](CONTRACTS.md).

## Context

The original design put a stateless **Surface Gateway** (Cloud Run) between the add-in and
Gemini Enterprise to hold Google credentials, federate identity, screen with Model Armor, route
to agents, and relay the stream. Research into the live **Discovery Engine `v1alpha`** API and
**Workforce Identity Federation** shows most of that is either unnecessary or already handled by
Google-side configuration:

1. **Identity federation can happen in the browser.** The STS token-exchange endpoint
   (`https://sts.googleapis.com/v1/token`, RFC 8693) accepts the user's Entra **OIDC token** as
   the `subjectToken` and returns a short-lived Google access token, directly from a browser/
   webview — no service-account key, no backend. (Google's own "WIF with API-based web apps"
   guide demonstrates the direct-from-browser flow.)
2. **Model Armor, agent routing, and grounding data stores are engine/assistant config**, not
   request parameters. In `v1alpha`, `StreamAssistRequest` has **no `agentsSpec`** — the
   assistant is configured in Gemini Enterprise, and the client just calls `:streamAssist`.
   (This also side-steps the `agentsSpec` agent-id bug.)
3. **`streamAssist` is already the server-side relay.** It streams grounded tokens + citations.
   Re-relaying through our own service adds latency and a credential-handling liability without
   adding capability.

## Decision

Build the product as an **Office.js add-in that calls Discovery Engine directly** as the
signed-in user. The thin, surface-agnostic `@ge/gemini-client` package owns:

- `WifTokenClient` — Entra token → STS exchange → cached Google access token (TTL-aware).
- `StreamAssistClient` — build the `StreamAssistRequest` from a `UnitDescriptor`, POST to
  `:streamAssist`, parse the chunked response into the `@ge/contracts` `SseEvent` stream
  (tokens + citations + provenance), resume via session id.
- Optional `proxyUrl` — if a tenant blocks browser CORS to `discoveryengine.googleapis.com`
  or wants a single audited egress point, the same client posts to a thin pass-through proxy
  instead. This is the **only** reason to run server code, and it stays optional.

Guardrails, agents, and grounding stores are **configured in the Gemini Enterprise engine** and
are out of scope for this repo (the admin's responsibility).

## Consequences

- **`services/gateway` becomes optional** (the CORS/audit proxy), not the centerpiece. Phase 0
  tasks 0.3–0.9 are reframed: federation + relay move client-side into `@ge/gemini-client`;
  Model Armor/provenance-signing are Google-side or client-side metadata writes.
- **No long-lived Google secrets in any client** — the client only ever holds the user's
  short-lived Entra token and the federated Google access token derived from it, in memory.
- **Packaging** targets the **unified Microsoft 365 manifest** (the modern successor to the
  standalone add-in manifest): one app package spanning Word/Excel/PowerPoint/**Outlook** task
  panes + Teams, and optionally a **Copilot declarative agent** as a natural-language entry point.
- **Surfaces are the work.** Effort shifts from backend plumbing to getting each surface's UX
  right against `docs/mockups/*`. **Outlook** is added as a first-class surface.

## What stays from the original design
The `@ge/contracts` boundary (UnitDescriptor, Finding, SSE events, Provenance), content-anchored
findings (`body.search` + re-resolve), provenance in host metadata, reversible/tracked writes,
identity-scoped reads, and residency pinning (now the regional `discoveryengine` endpoint chosen
client-side).
