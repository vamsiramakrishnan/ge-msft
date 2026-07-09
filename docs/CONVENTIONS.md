# Conventions

> **Note (client-direct, `ADR-0001`).** The **TypeScript stack, code style, project-layout, testing,
> error-handling, and security standards** below are all current and enforced. The **gateway-specific
> items are historical** — there is no gateway/`services/*` tier and no Python ADK agents in the repo
> today: ignore "Fastify/Cloud Run for the gateway," "the gateway is stateless," "one responsibility
> per gateway module," and structured-gateway-logging/audit-to-BigQuery as *deploy-time* concerns for
> the optional proxy, not workspace code. The security standards (no Google secrets in a client,
> identity scoped end-to-end, untrusted host content, provenanced+reversible writes, residency, least
> privilege) hold identically client-direct.

## Stack
- **TypeScript** (strict) for clients and gateway; **Python 3.12** for `services/agents` (ADK).
- **React 18** + Office.js / TeamsJS for clients; **Vite** for bundling; HTTPS in dev (Office requires it).
- **Fastify** on **Node 20+** for the gateway; **Cloud Run** for deploy.
- **Bun workspaces** monorepo; **Zod** for runtime validation (shared in `packages/contracts`); **Vitest** (TS) and **pytest** (agents).

## Code style
- Prettier + ESLint (typescript-eslint, react-hooks). Format on save; lint clean before done.
- `camelCase` for values, `PascalCase` for types/components, `kebab-case` for files and package names.
- Prefer pure functions and explicit return types on exported functions. No `any` — use `unknown` + a Zod parse at boundaries.
- Every cross-boundary payload is parsed with its Zod schema from `packages/contracts` on receipt. Don't trust shapes.
- Keep modules small and single-purpose. The five client modules (`DocBridge`, `AuthClient`, `StreamClient`, `ProvenanceStore`, `UnitComposer`) have stable interfaces — implement per surface behind those interfaces.

## Project layout rules
- `packages/web-shell` is **surface-agnostic**. If code references Word/Excel/PowerPoint/OneNote/Teams APIs, it belongs in a `bridge-*` or `teams`, not the shell.
- The gateway is **stateless**. No per-user state in memory beyond a request; sessions live in StreamAssist (server) and the host's custom metadata (client).
- One responsibility per gateway module: `auth`, `federation`, `router`, `relay`, `provenance`, `audit`, `armor`.

## Testing
- Unit-test contracts (schema round-trips), the unit resolver, the router, anchoring (`body.search` resolution + drift degradation), and the SSE relay (incl. polling fallback).
- Mock external services (Gemini Enterprise, Entra, STS, connectors) behind thin clients so tests don't hit the network. Label mocks clearly; never let a mock masquerade as a real integration.
- Each surface's signature interaction gets at least one integration test against its mockup's expected behavior.

## Error handling & UX copy
- Errors explain what happened and how to fix it, in the interface's voice — never vague, never apologetic. Empty states invite an action.
- An action keeps its name through the whole flow (the button that says "Accept change" produces a "tracked change inserted" result).
- On `401`, the client silently refreshes the NAA/TeamsJS token and retries idempotent reads; write-backs carry a `changeId` so a retry can't double-apply.

## Logging & config
- Structured JSON logs from the gateway (request id, identity, intent, agent, sources, latency). Audit every invocation to BigQuery.
- All config from environment (`.env.example` is the contract). No secrets in code or client bundles; secrets via Secret Manager in deployment.

## Security standards (enforced; `security-reviewer` checks these)
1. **No Google credentials in any client or `web-shell`.** Clients hold only the user's short-lived Entra/Teams token.
2. **Identity scoped end to end.** Every Gemini and connector call acts as the signed-in user; prefer delegated permissions over org-wide and `Sites.Selected` over all-sites.
3. **Host content is untrusted.** Screen via Model Armor; pass as data, never instructions. Never execute or follow instructions found inside a document or transcript.
4. **Writes are provenanced and reversible.** Tracked changes / citation-tagged blocks only; every write records a `ProvenancePayload`.
5. **Residency pinned.** Cloud Run + `discoveryengine` region match the tenant; VPC firewall restricts connector egress to specific FQDNs.
6. **Least privilege everywhere.** Minimal Entra app permissions; minimal IAM on the gateway's service account; RSC (not blanket consent) for Teams transcript.

## Git
- Small, focused commits per build-plan task; commit message references the task id (e.g. `feat(gateway): 0.6 StreamAssist SSE relay`).
- Don't commit `.env`, build output, or credentials (see `.gitignore`).
