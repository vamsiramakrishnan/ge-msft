# Bug report: `AgentService.ListAgents` returns 403 for a WIF caller that *holds* `discoveryengine.agents.list`

> **RESOLUTION / WORKAROUND (2026-07-08):** for our use case the working call is
> **`POST …/assistants/{a}:listAvailableAgentViews`** (custom verb → `discoveryengine.agents.listAvailableAgentViews`,
> which the WIF principal holds) — it returns **200** with all available agents (owned + enabled managed)
> from the same WIF token. `AgentService.ListAgents` below still 403s and the contradiction is likely a
> real limitation/bug (permission held, method denied), but it no longer blocks us. Question for the DE
> team stands: should `ListAgents` succeed when `discoveryengine.agents.list` is held?


## Summary

A caller authenticated via **Workforce Identity Federation** (Entra → Google STS) is granted the IAM
permission `discoveryengine.agents.list` (confirmed by `projects.testIamPermissions`), but
`AgentService.ListAgents` (`GET …/assistants/*/agents`) returns **403 PERMISSION_DENIED — "User does not
have permission to list all of the agents."** On the *same* assistant resource, `AgentService.GetAgent`
(`GET …/agents/{id}`) succeeds with **200**. The 403 is independent of `filter` / `pageSize` / `orderBy`.

This looks like an **application-level authorization check inside Discovery Engine that is not backed by
the `discoveryengine.agents.list` IAM permission** — or that resolves the caller's identity differently
for WIF principals than for browser/OAuth users.

## Environment / resources

- Project: `saib-ai-playground` (number `288406675721`), org `620724670057`
- Engine: `ge-msft-plugin-test_1782382759735`, collection `default_collection`, assistant `default_assistant`, location `global`
- API: `discoveryengine.googleapis.com/v1alpha`
- Caller identity (Workforce Identity Federation):
  `principal://iam.googleapis.com/locations/global/workforcePools/saib-wf-pool/subject/vamramak@psott.onmicrosoft.com`
- Token: `gcloud auth print-access-token` after `gcloud auth login --login-config <workforce-login-config>`
  (a workforce-federated access token; `$TOK` below). Header `X-Goog-User-Project: saib-ai-playground` on every call.
- The skill agents in this assistant are `skillAgentDefinition` agents, `state: PRIVATE`, owner
  `vamramak@psott.onmicrosoft.com` (as shown by the widget `widgetListAvailableAgentViews` view).

## The contradiction (this is the crux)

**(a) IAM says the permission is held** — `projects.testIamPermissions`:

```bash
curl -s -X POST \
  "https://cloudresourcemanager.googleapis.com/v1/projects/saib-ai-playground:testIamPermissions" \
  -H "Authorization: Bearer $TOK" -H "content-type: application/json" \
  -d '{"permissions":["discoveryengine.agents.list","discoveryengine.agents.get",
       "discoveryengine.agents.listAvailableAgentViews","discoveryengine.agents.getAgentView",
       "discoveryengine.assistants.assist","discoveryengine.dataConnectors.acquireAccessToken"]}'
```
→ response includes **`discoveryengine.agents.list`** (and `agents.get`, `agents.listAvailableAgentViews`,
`agents.getAgentView`, `assistants.assist`, `dataConnectors.acquireAccessToken`) in the granted set.
The held permission set matches `roles/discoveryengine.user`.

**(b) The REST list call is nonetheless denied:**

```bash
curl -s -o /dev/null -w "%{http_code}\n" \
  -H "Authorization: Bearer $TOK" -H "X-Goog-User-Project: saib-ai-playground" \
  "https://discoveryengine.googleapis.com/v1alpha/projects/288406675721/locations/global/collections/default_collection/engines/ge-msft-plugin-test_1782382759735/assistants/default_assistant/agents"
```
→ **HTTP 403**
```json
{"error":{"code":403,"status":"PERMISSION_DENIED",
  "message":"User does not have permission to list all of the agents."}}
```

Same 403 with query parameters (all tested):
```
…/agents?filter=display_name=%22m365-surface-commander%22   → 403
…/agents?pageSize=1                                          → 403
…/agents?pageSize=50&orderBy=update_time%20desc             → 403
```

**(c) GetAgent on the same resource works (200):**

```bash
curl -s -o /dev/null -w "%{http_code}\n" \
  -H "Authorization: Bearer $TOK" -H "X-Goog-User-Project: saib-ai-playground" \
  "https://discoveryengine.googleapis.com/v1alpha/projects/288406675721/locations/global/collections/default_collection/engines/ge-msft-plugin-test_1782382759735/assistants/default_assistant/agents/3708891467397998816"
```
→ **HTTP 200**, returns the full `Agent` with `skillAgentDefinition` (instruction, gcsUri, owner, subfiles),
`displayName: m365-surface-commander`, `state: PRIVATE`.

## Everything else we verified (same WIF token)

| Call | Result |
|---|---|
| `assistants.agents.get` (by id) | **200** |
| `assistants.agents.create / patch / delete` | **200** (create/update/delete a throwaway skill worked) |
| raw `…/agents/{id}/files:upload?upload_protocol=raw` (zip) | **200** (SKILL.md→instruction, rest→subfiles) |
| `agents/{id}/a2a/v1/card` (getCard) | **200** |
| `assistants:streamAssist` | **200** (skill routes via `mention://?uri=<id>` in query text) |
| **`assistants.agents.list`** | **403** "list all of the agents" ← this report |
| `assistants.get` | **403** `discoveryengine.assistants.get` denied |
| `assistants.list` | **403** `discoveryengine.assistants.list` denied |
| `agents:deploy` | **403** `discoveryengine.agents.deploy` denied |
| `content-discoveryengine…:widgetAcquireAccessToken` (WIF bearer) | **401** "Authentication failed" |
| `content-discoveryengine…:widgetListAvailableAgentViews` (short-lived **widget JWT** from an authenticated browser session, `iss=vertexaisearch.cloud.google`, `aud=content-discoveryengine`) | **200** — lists the same agents fine |

So the identical agents **are** listable — but only via the widget endpoint with a browser/widget JWT,
not via the public `AgentService.ListAgents` with the WIF token, even though that token holds
`discoveryengine.agents.list`.

## Hypothesis / question for the DE team

1. Does `AgentService.ListAgents` enforce an **application-level owner/identity check** ("agents created
   by the caller") *in addition to* the `discoveryengine.agents.list` IAM permission? The method is
   documented as "Lists all Agents under an Assistant **which were created by the caller**."
2. If so, how is "the caller/creator" identity resolved for a **Workforce Identity Federation** principal
   (`principal://…/workforcePools/…/subject/<entra-upn>`)? The agents' `owner` is `vamramak@psott.onmicrosoft.com`,
   which equals the WIF `subject`, yet ListAgents still denies — suggesting GE matches the creator against
   a **Google/widget identity** (the browser widget's `sub` is `csesidx/<n>`), not the WIF subject string.
3. Is there a supported way to make `ListAgents` succeed for a WIF caller (a role/permission, a resource
   condition, or a filter that scopes to the caller), or is listing intentionally only available to the
   Google/OAuth/widget identity that owns the agents?

Expected: with `discoveryengine.agents.list` held, `ListAgents` returns the caller's own agents (200),
consistent with `GetAgent` succeeding on those same agents.

## Repro (minimal)

```bash
export CLOUDSDK_CONFIG=<repo>/.gcloud
gcloud auth login --login-config <repo>/.gcloud/saib-wif-login-config.json --no-launch-browser
TOK=$(gcloud auth print-access-token)
BASE=https://discoveryengine.googleapis.com/v1alpha/projects/288406675721/locations/global/collections/default_collection/engines/ge-msft-plugin-test_1782382759735/assistants/default_assistant
H=(-H "Authorization: Bearer $TOK" -H "X-Goog-User-Project: saib-ai-playground")

curl -s -w '\n%{http_code}\n' "${H[@]}" "$BASE/agents/3708891467397998816"   # 200
curl -s -w '\n%{http_code}\n' "${H[@]}" "$BASE/agents"                        # 403 "list all of the agents"
```
