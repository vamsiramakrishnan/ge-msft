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

**Method references**
- `streamAssist.md` — primary streaming grounded assistant (chat over the unit).
- `search.md` — Vertex AI Search (retrieval, snippets, summaries, extractive answers).
- `answer.md` — single-shot grounded answer with citations + related questions.
- `sessions.md` — multi-turn conversation state (resume across surfaces/sessions).
- `autocomplete.md` — `completeQuery` suggestions for the composer.
- `grounding-check.md` — verify a rewrite is grounded before applying it.
- `ranking.md` — semantic reranking of candidates.
- `assistant-config.md` — engine/assistant config (Model Armor, tools, canned queries) — read-only reference.
- `methods-index.md` — every method in the surface (index).

**How-to / mechanisms** (research distilled for the add-in)
- `context-mechanisms.md` — how to attach context (`query.parts[]`, reference-over-inline) and scope
  connectors (`toolsSpec.dataStoreSpecs` + `filter` + `boost`), `actionSpec`, sessions.
- `agent-invocation.md` — targeting a specific agent (no `agentsSpec` in `v1alpha`; assistant
  targeting + `actionSpec`; A2A-direct caveats). **Superseded in part by `skills-and-agents.md`.**
- `skills-and-agents.md` — the **verified** skill lifecycle: skills are `agents` with a
  `skillAgentDefinition`, created/uploaded/shared on the live `v1alpha` endpoint and **mounted
  per-turn via `skillsSpec`** (the field is real on the wire though absent from the published
  schema). Where the `/` + `@` command skills (`skill/`) plug in, plus the skill↔workspace parity
  tasks.
- `a2ui.md` — A2UI agent-authored interactive UI and how the add-in maps its actions to host
  actuations.
- `files-and-limits.md` — inline context vs session context files (`addContextFile`), **code
  execution** for xlsx/csv analysis, the verified quotas, and the inline/reference/upload decision
  policy.

## Endpoint & regions
- Global: `https://discoveryengine.googleapis.com`
- Regional (residency): `https://discoveryengine.<region>.rep.googleapis.com` (e.g. `eu`, `us`,
  `asia-northeast1`). Pin to the tenant's residency commitment.

## Auth
`Authorization: Bearer <google-access-token>` where the token is obtained from STS token
exchange (`https://sts.googleapis.com/v1/token`, grant_type `token-exchange`, subject =
the Entra OIDC token, audience = the Workforce Identity Pool **provider** URI). No Google
service-account key is ever held by the client. See `packages/gemini-client`.
