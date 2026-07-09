# Listing & reading agents/skills — what works with the WIF token

How the add-in enumerates and reads Gemini Enterprise **skill agents** (create → `skillAgentDefinition`).
Written from live probes against the **saib** tenant (`saib-ai-playground`, engine
`ge-msft-plugin-test_1782382759735`, principal `vamramak@psott.onmicrosoft.com`, WIF token) plus the
public `v1alpha` discovery document. Each row is tagged **VERIFIED** (live-probed) or **DOCUMENTED**
(from the discovery doc / GE reference, not yet live-tested here).

## ★ THE WORKING LIST (VERIFIED 2026-07-08): `assistants:listAvailableAgentViews`

`agents.list` is the wrong method for us (app-gated "list all created by caller" → 403 even though we
hold the IAM permission). The method that **works with the plain WIF token** is the custom verb
**`POST …/assistants/{a}:listAvailableAgentViews`** — "agents *available to me*", mapping to
`discoveryengine.agents.listAvailableAgentViews` (which the WIF principal holds). Public endpoint, no
widget JWT, no extra grant.

```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOK" -H "X-Goog-User-Project: saib-ai-playground" \
  -H "Content-Type: application/json" -d '{}' \
  "https://discoveryengine.googleapis.com/v1alpha/projects/288406675721/locations/global/collections/default_collection/engines/ge-msft-plugin-test_1782382759735/assistants/default_assistant:listAvailableAgentViews"
```
→ **HTTP 200**, `{ "agentViews": [ … ] }`, returning **all 4** agents in the engine:

| id | displayName | state | agentType |
|---|---|---|---|
| `3708891467397998816` | m365-surface-commander | PRIVATE | SKILL_AGENT |
| `m365-command-planner` | m365-command-planner | PRIVATE | SKILL_AGENT |
| `deep_research` | Deep Research | ENABLED | MANAGED |
| `default_idea_generation` | Idea Generation | ENABLED | MANAGED |

Each `agentView` carries: `name` (full public resource name — id = last segment), `displayName`,
`description`, `state`, `agentType` (`SKILL_AGENT` \| `MANAGED`), `agentOrigin`, `agentSharingState`,
`sharingConfig`, `ownerUserPrincipal`, `userPermissions`, `dataStoreSpecs`, `toolSelection`,
`skillAgentInfo`, `updateTime`. It's **POST-only** (GET → 404); body accepts `pageSize`, `pageToken`,
and (per the v1main proto) `filter`/`sortBy`. This is the list path the taskpane should use.

(Note it also revealed the planner is back as the **slug id `m365-command-planner`**, not the tombstoned
numeric `10348…` — and it surfaces Google **MANAGED** agents like Deep Research that `agents.list` would
never return, since those weren't "created by the caller".)

## Data stores / connectors / engines — same get✓/list✗ split, and the grounding implication

Verified 2026-07-08 (WIF token + `testIamPermissions`): the "**`.get` by id works, `.list` denied**"
pattern holds beyond agents.

| Call | Result |
|---|---|
| `dataStores.list` (`GET …/collections/{c}/dataStores`) | **403** — `discoveryengine.dataStores.list` not held |
| `engines.list` (`GET …/collections/{c}/engines`) | **403** — `discoveryengine.engines.list` not held |
| `dataConnector.get` (`GET …/collections/{c}/dataConnector`) | **403** — `discoveryengine.dataConnectors.get` not held |
| `dataStores.get` / `engines.get` (by id) | **held ✓** (read a known resource) |
| `servingConfigs.search` | **held ✓** (search a known data store) |
| `dataConnectors.queryAvailableActions` | **held ✓** (list a connector's actions) |
| `dataConnectors.acquireAccessToken` | **held ✓** |
| `collections.get` | **held ✓** |

(The claim that `dataStores.list`/`engines.list`/`dataConnectors.get` are in `roles/discoveryengine.user`
is incorrect — that role carries `dataStores.get`/`engines.get` (by id) + search/action perms, not the
list/get-config verbs.)

**Implication (multi-source grounding — achievable today):** we can't *discover* connectors/data stores
(no `.list`), but grounding doesn't need discovery. Pass known **data-store ids** in
`toolsSpec.vertexAiSearchSpec.dataStoreSpecs[]` and streamAssist grounds on them (SharePoint, Salesforce,
OneDrive, …) — `gemini-client` already maps `config.dataStores` → `dataStoreSpecs`. `dataStores.get`
validates each id; `servingConfigs.search` retrieves directly.

### Selecting data stores — VERIFIED recipe (no `dataStores.list` needed)

`dataStores.list`/`collections.list` are denied, but **`engines.get` is held** and the **Engine resource
carries `dataStoreIds[]`** — the stores attached to the app. That is our discovery path:

```
1. GET …/engines/{engine}                     (engines.get ✓)  → engine.dataStoreIds[]
2. GET …/collections/{c}/dataStores/{id}       (dataStores.get ✓) → displayName/vertical, per id, for a picker
3. streamAssist toolsSpec.vertexAiSearchSpec.dataStoreSpecs:[{ dataStore: ".../dataStores/<id>" }]  → ground
   (servingConfigs.search ✓ / .answer ✓ also available for direct retrieval)
```

Live (2026-07-08) our engine `ge-msft-plugin-test_…` returns **10 federated stores** via `dataStoreIds`:
`msft-outlook-fed_…_{mail, mail-attachment, calendar, contact}`, `msft-onedrive-fed_…_file`,
`msft-sharepoint-fed_…_{attachment, comment, event, page, file}` — i.e. **Outlook + OneDrive + SharePoint
are already connected**. `collections.get` returns only `{name, createTime}` (no data-store info);
`dataConnectors.get` is denied — so `engines.get` is the one source of truth for "which data stores can I
ground on." The taskpane's data-store picker = read `dataStoreIds`, group by connector prefix, user
selects, pass into `dataStoreSpecs`.

## Wired in code (2026-07-08)

- `gemini-client`: `listAvailableAgentViews()` / `listSkillAgents()` (POST `:listAvailableAgentViews`),
  `getAgent()` (by id), `listEngineDataStores()` / `getEngineDataStoreIds()` / `getDataStore()`
  (`engines.get` → `dataStoreIds` → per-store), `dataStoreSpecsFromIds()` (build grounding), plus the
  URL/resource helpers in `config.ts` (`listAvailableAgentViewsUrl`, `engineUrl`, `dataStoreUrl`,
  `dataStoreResourceName`). All unit-tested with a mock fetch.
- `web-shell` `composeSession({ discoverCatalog: true })` runs both discoveries at boot (best-effort,
  WIF-authenticated) and returns `availableAgents` + `availableDataStores` in `ComposedSession`; `main.tsx`
  passes `discoverCatalog: true` and logs the counts/connectors. These are the sources a skill picker and
  a data-store picker render from; selected data-store ids feed `dataStoreSpecs` for grounding.

## TL;DR

- To **read one known skill** (planner/commander by id): **`agents.get`** — public, WIF-authenticated, **VERIFIED working**. No widget, no extra IAM.
- To **list all skills**: the public method is **`agents.list`**, but it is **VERIFIED 403** for our
  principal (needs the IAM permission `discoveryengine.agents.list`). Grant `roles/discoveryengine.viewer`
  and it works with the same WIF token.
- **`engines.assistants.list`** (the endpoint linked in discussion) lists **assistants, not agents** —
  useful to discover the assistant path, not to fetch skills.
- The **widget** `widgetListAvailableAgentViews` lists skills **without** the IAM grant, but needs a
  short-lived **widget JWT** (`iss=vertexaisearch.cloud.google`, `aud=content-discoveryengine`, ~5 min)
  — not the WIF token.

## The method matrix

All public REST is `https://discoveryengine.googleapis.com/v1alpha/…`. OAuth scope for every row below
is `cloud-platform` / `discoveryengine.assist.readwrite` / `.readwrite` / `.serving.readwrite` (the WIF
token carries `cloud-platform`), so **differences are IAM permission, not scope**.

| Goal | Method | Endpoint | Auth needed | Status |
|---|---|---|---|---|
| Read one skill by id | `assistants.agents.get` | `GET …/assistants/{a}/agents/{id}` | WIF | **VERIFIED ✓** — returns full `skillAgentDefinition` (instruction, gcsUri, owner, subfiles) |
| List **all** skills (public) | `assistants.agents.list` | `GET …/assistants/{a}/agents` | WIF **+ `discoveryengine.agents.list` IAM** | **VERIFIED 403** without grant ("does not have permission to list all of the agents"). Lists agents *created by the caller*; params `filter` (display_name), `orderBy`, `pageSize`, `pageToken` |
| List **assistants** (not agents) | `engines.assistants.list` | `GET …/engines/{e}/assistants` | WIF **+ `discoveryengine.assistants.list` IAM** | **VERIFIED 403** (`discoveryengine.assistants.list` denied). Even if granted it returns `ListAssistantsResponse {assistants[]}` — assistants, **not** agents/skills. Not a "fetch all agents" path |
| Read the assistant | `engines.assistants.get` | `GET …/assistants/{a}` | WIF **+ `discoveryengine.assistants.get` IAM** | **VERIFIED 403** (`discoveryengine.assistants.get` denied). So "get-assistant-then-list" is not a viable chain — the assistant read itself is denied. The `Assistant` resource also carries **no agent list** |
| Read a skill's capabilities | `assistants.agents.a2a.v1/getCard` | `GET …/agents/{id}/a2a/v1/card` (stable **v1**) | WIF | **VERIFIED ✓** — returns the A2A AgentCard even for skill agents |
| List **all** skills (widget) | `widgetListAvailableAgentViews` | `POST content-discoveryengine…/locations/{loc}/widgetListAvailableAgentViews` | **widget JWT** + `x-server-token` (not WIF) | **VERIFIED ✓** with a widget JWT — returns `agentViews[] {name, displayName, state, userPermissions, …}`, filter `agent_type = SKILL_AGENT`, `agentOrigin: USER|GOOGLE` |

## ⚠ CORRECTION (live-verified): `agents.list` is blocked at the APP layer, not IAM

Later probing overturned the "just grant the list permission" theory. `POST projects/…:testIamPermissions`
reports the WIF principal **HOLDS** `discoveryengine.agents.list`, `agents.listAvailableAgentViews`,
`agents.getAgentView`, `agents.{get,create,update,delete}`, `assistants.assist`,
`dataConnectors.acquireAccessToken`, `sessions.create`, `notebooks.list`, `projects.get` — i.e. the
full **`roles/discoveryengine.user`** set (it lacks only `assistants.get/list`, `agents.deploy`,
`projects.getIamPolicy`). **Yet `agents.list` still returns 403 "does not have permission to list all of
the agents"** — with or without a `filter`/`pageSize`. So:

- The 403 is **application-level authorization inside Discovery Engine**, gated *separately* from the
  IAM permission. Holding `discoveryengine.agents.list` does **not** satisfy it.
- The likely cause: GE keys "agents created by the caller" to the **Google/widget identity**, not the
  **WIF workforce `subject`** — so the federated principal isn't recognized as the owner, even though
  the email (`vamramak@…`) matches. A `filter` does not help (tested).
- **Consequence:** granting `roles/discoveryengine.user`/`viewer` will NOT make `agents.list` work from
  the WIF identity. Listing all agents is only reachable via the **widget** (browser/OAuth identity) or
  the internal path — not the federated principal.
- **What still works from WIF:** `agents.get` by id, `streamAssist`, `getCard`, create/update/delete.

Net for the taskpane: **use `agents.get`-by-id** (works). Treat full list-all as unavailable to the WIF
identity on this tenant.

## Why `agents.get` works but `agents.list` 403s (original IAM framing — superseded above)

Same OAuth scopes; different **IAM permissions**. Our principal has *assist* + *agent read/write on
resources it owns* (get/create/update/delete all VERIFIED working), but **not** the list-all permission
`discoveryengine.agents.list`. So:

- **get-by-id** (`discoveryengine.agents.get`) → granted → works.
- **list-all** (`discoveryengine.agents.list`) → not granted → 403.

The fix is a one-time IAM grant. **The right role is `roles/discoveryengine.user`** (user-level, not
admin) — it contains `discoveryengine.agents.list` **and** the full user-level set the add-in already
uses: `agents.{get,create,update,delete,getAgentView,listAvailableAgentViews}`, `assistants.assist`
(streamAssist), `sessions.*`, `notebooks.*`, and `dataConnectors.acquireAccessToken` (the widget-token
mint). Our WIF principal currently holds only a *subset* (assist + `agents.get`), which is exactly why
get works and list 403s.

| Denied | Grant |
|---|---|
| `discoveryengine.agents.list` (list skills) + full assist/agent user set | **`roles/discoveryengine.user`** (preferred — the signed-in-user role) |
| read-only alternative | `roles/discoveryengine.viewer` |
| `discoveryengine.agents.deploy` (managed/app agents only — skills don't need it) | agent-admin role |
| sharing (`sharingConfig` PATCH) | agent-admin role |

## "But the docs say it lists agents created by the caller — why 403?" (identity type)

`agents.list` is described as *"Lists all Agents under an Assistant which were created by the caller."*
That is true — for a caller who holds `discoveryengine.agents.list`. The public
[Medium walkthrough](https://medium.com/google-cloud/integrating-agents-on-custom-platforms-with-gemini-enterprise-04f8bb3b52ca)
that lists agents successfully uses **end-user OAuth 2.0** (a consent-screen OAuth Client; 3-legged) —
*"all API calls must be authenticated with the user's credentials"* — and that user carries a viewer/editor
role, so the permission is present.

**We authenticate differently.** Our token is a **Workforce Identity Federation** principal
(`principal://iam.googleapis.com/locations/global/workforcePools/saib-wf-pool/subject/vamramak@…`), which
is a distinct identity from a normal Google IAM user and has **not** been granted
`discoveryengine.agents.list`. `agents.get` on the *same* `default_assistant` path returns **200** (that
permission is present), so the assistant id and endpoint are correct — only *list-all* is missing. VERIFIED.

**Grant it to the workforce principal (or the whole pool):**
```
gcloud projects add-iam-policy-binding saib-ai-playground \
  --member="principal://iam.googleapis.com/locations/global/workforcePools/saib-wf-pool/subject/vamramak@psott.onmicrosoft.com" \
  --role="roles/discoveryengine.viewer"
# or the whole pool:
#   --member="principalSet://iam.googleapis.com/locations/global/workforcePools/saib-wf-pool/*"
```
After this, public `agents.list` works with the WIF token — no widget, no end-user OAuth. (The
alternative — adopting 3-legged end-user OAuth like the article — diverges from our client-direct WIF
model in ADR-0001 and is not recommended.)

**VERIFIED: the WIF principal cannot self-grant this.** Running the binding above *as* the workforce
principal fails: `does not have permission to access projects instance [saib-ai-playground:getIamPolicy]`
— it lacks `resourcemanager.projects.setIamPolicy`. So the grant must be performed **out of band by a
GCP project IAM admin** (`roles/resourcemanager.projectIamAdmin`/owner), via Console or an admin
identity. Nothing the add-in or its WIF token can do at runtime fixes list-all; treat full discovery as
an admin-provisioned capability, and default the taskpane to **`agents.get` by id** (which needs no grant).

## Recommended approach for the taskpane (dynamic routing)

Two tiers, both public (`discoveryengine`, **not** the widget), WIF-authenticated:

1. **Known skills → `agents.get` by id (works today, no grant).** The add-in already knows the
   planner/commander ids from `.env`; resolve each with `getAgent(cfg, id, {tokens})`
   (`packages/gemini-client/src/agents.ts`) to confirm existence + read `displayName`/`state`, then
   configure routing. Falls back to `.env` only on a 404. This is the default boot path.
2. **Full discovery → `agents.list` (needs `roles/discoveryengine.viewer`).** Once granted,
   `DiscoveryCatalogClient.listAdminSkills()` (already calls public `agents.list`) enumerates every
   skill with the WIF token. Enable this by preferring the public path over the widget branch (today
   the catalog prefers widget whenever `widget.configId` is set — and it sends the **WIF** token to the
   widget endpoint, which **401s**, silently falling back to `.env`; that branch should be gated behind
   an explicit "widget-only tenant" flag).

**Widget listing is the fallback only** — for tenants where the IAM grant isn't available — and requires
minting a widget JWT (see `widget-service-skills.md` / `skill/acquire_widget_token.py`), which the WIF
token is not.

### Can we mint the widget JWT programmatically (no browser)? — VERIFIED: not with WIF

`WidgetService.WidgetAcquireAccessToken`
(`POST content-discoveryengine…/locations/{loc}/widgetAcquireAccessToken`, returns `uToken`) is the
programmatic mint the browser widget uses. Live-tested on saib:
- **WIF Bearer → `401 Authentication failed`** (the widget host rejects the workforce principal).
- **No auth (server-token + configId only) → `401 … Expected OAuth 2 access token, login cookie or other valid credential`.**

The widget JWT's `sub` is a browser client-session id (`csesidx/…`), not our federated identity. So the
acquire requires a **Google end-user OAuth 2.0 token or a browser login cookie** — the WIF workforce
token cannot mint it. Two real programmatic options, neither of which is "WIF drives the widget":
1. **End-user OAuth (3LO):** a consent-screen OAuth client + a stored user refresh token (the
   [Medium](https://medium.com/google-cloud/integrating-agents-on-custom-platforms-with-gemini-enterprise-04f8bb3b52ca)
   model). That token *can* call `widgetAcquireAccessToken` and the public `agents.list`. But it
   reintroduces user consent + refresh-token custody, diverging from client-direct WIF (ADR-0001).
2. **Skip the widget:** admin grants `roles/discoveryengine.viewer` to the WIF pool → public
   `agents.list` works directly with the WIF token we already have. **Recommended** — one IAM binding
   replaces the entire widget-token dance.

## Provenance of these findings

- `agents.get` ✓, `agents.list` 403, `create/update/delete` ✓, `getCard` ✓, bundle upload ✓ — live-probed
  on saib (see [[ge-agent-invocation-live-findings]] memory and `docs/api/discoveryengine/skills-and-agents.md`).
- `engines.assistants.list` — **live-tested on saib: 403** `discoveryengine.assistants.list` denied
  (same IAM gap as `agents.list`; shares the OAuth scope but the IAM permission is not granted). Shape
  from the discovery doc: `ListAssistantsResponse {assistants[]}` — lists assistants, not agents.
- `widgetListAvailableAgentViews` ✓ — the authenticated widget cURL captured from the GE web UI
  (returns `m365-surface-commander`, `state: PRIVATE`; the planner id was absent = tombstoned).
