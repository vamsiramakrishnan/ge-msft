# Discovery Engine / Gemini Enterprise API — knowledge base

Vendored, distilled reference for the **Discovery Engine API `v1alpha`** (the REST surface
behind Gemini Enterprise), captured for this add-in. Source of truth:

- Discovery document: `https://discoveryengine.googleapis.com/$discovery/rest?version=v1alpha`
- Reference docs: https://docs.cloud.google.com/gemini/enterprise/docs/reference/rest

The add-in is **client-direct**: the Office.js task pane exchanges the signed-in user's Entra
token for a Google access token via Workforce Identity Federation (STS), then calls these
endpoints directly as the user. Guardrails (Model Armor), agent routing, and grounding data
stores are configured **in the Gemini Enterprise engine/assistant**, not in our code — see
`assistant-config.md`.

## Files
- `streamAssist.md` — primary streaming grounded assistant (chat over the unit).
- `search.md` — Vertex AI Search (retrieval, snippets, summaries, extractive answers).
- `answer.md` — single-shot grounded answer with citations + related questions.
- `sessions.md` — multi-turn conversation state (resume across surfaces/sessions).
- `autocomplete.md` — `completeQuery` suggestions for the composer.
- `grounding-check.md` — verify a rewrite is grounded before applying it.
- `ranking.md` — semantic reranking of candidates.
- `assistant-config.md` — engine/assistant config (Model Armor, tools, canned queries) — read-only reference.
- `methods-index.md` — every method in the surface (index).

## Endpoint & regions
- Global: `https://discoveryengine.googleapis.com`
- Regional (residency): `https://discoveryengine.<region>.rep.googleapis.com` (e.g. `eu`, `us`,
  `asia-northeast1`). Pin to the tenant's residency commitment.

## Auth
`Authorization: Bearer <google-access-token>` where the token is obtained from STS token
exchange (`https://sts.googleapis.com/v1/token`, grant_type `token-exchange`, subject =
the Entra OIDC token, audience = the Workforce Identity Pool **provider** URI). No Google
service-account key is ever held by the client. See `packages/gemini-client`.
