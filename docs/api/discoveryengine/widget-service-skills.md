# WidgetService skill lifecycle and send-time skill references

Source reference:

- Official RPC reference:
  <https://docs.cloud.google.com/generative-ai-app-builder/docs/reference/rpc/google.cloud.discoveryengine.v1alpha#widgetservice>
- The official `WidgetService` page labels these methods as internal and subject to breaking
  changes. Treat the widget path as a developer/private-skill path, not a CI or tenant-admin API.
- The public RPC page documents the WidgetService family, but the private skill-agent endpoints
  below (`widgetCreateAgent`, `widgetListAvailableAgentViews`, `widgetDeleteAgent`, raw
  `files:upload`) are observed from the Gemini Enterprise UI, not a durable public admin contract.

This repo uses two Discovery Engine surfaces:

1. **Public/admin API**: `discoveryengine.googleapis.com`, OAuth/ADC, IAM-scoped assistant-agent
   permissions.
2. **Widget API**: `content-discoveryengine.googleapis.com`, short-lived Vertex AI Search widget JWT
   from the signed-in Gemini Enterprise web session, widget permissions such as `canEdit` and
   `canDelete`.

For the current dev app, the reliable private-skill lifecycle is the widget path.

## Widget operation contract

Every widget call carries:

```json
{
  "configId": "<widget config GUID>",
  "additionalParams": {
    "token": "-",
    "origin": "ORIGIN_UNSPECIFIED"
  }
}
```

Required headers:

```text
Authorization: Bearer <short-lived widget JWT>
Origin: https://vertexaisearch.cloud.google
Referer: https://vertexaisearch.cloud.google/
x-server-token: <optional widget server token>
```

The token must be a Vertex AI Search widget JWT:

```text
iss = https://vertexaisearch.cloud.google
aud = https://content-discoveryengine.googleapis.com
```

Do not use ADC, `gcloud auth print-access-token`, cookies, SAPISIDHASH, or XSRF automation for
widget skill management.

## Skill lifecycle

| Operation | Widget endpoint | Payload shape |
| --- | --- | --- |
| List private skills | `POST /v1alpha/locations/{location}/widgetListAvailableAgentViews` | `listAvailableAgentViewsRequest: { pageSize: 200, filter: "agent_type = SKILL_AGENT", agentOrigin: "USER" }` |
| Create shell skill | `POST /v1alpha/locations/{location}/widgetCreateAgent` | `createAgentRequest: { agent: { displayName, description, skillAgentDefinition: { instruction } }, defaultFilesSkipped: true }` |
| Upload/update bundle | `POST /upload/v1alpha/{assistant}/agents/{agentName}/files:upload` | resumable upload: `start`, then `upload, finalize` |
| Verify | `POST /v1alpha/locations/{location}/widgetGetAgentView` | `getAgentViewRequest: { name: "<numeric agent name>" }` |
| Delete | `POST /v1alpha/locations/{location}/widgetDeleteAgent` | top-level `name: "<numeric agent name>"` |

There is no separate durable "update skill metadata and files" abstraction in our tooling. Updating
a bundle means uploading a new zip to the existing numeric agent, or deleting and recreating the
agent and then uploading the new zip.

## Local commands

Dry-run is the default.

```bash
python3 skill/update_skills.py --api-mode widget --list
python3 skill/update_skills.py --api-mode widget --live --list

python3 skill/update_skills.py --api-mode widget --live --upload-existing
python3 skill/update_skills.py --api-mode widget --live --replace --yes
python3 skill/update_skills.py --api-mode widget --live --delete-only --yes
python3 skill/update_skills.py --api-mode widget --live --create-new
```

For one skill:

```bash
python3 skill/create_skill.py --api-mode widget --live --list
python3 skill/create_skill.py --api-mode widget --live --upload-existing --agent-id 17573173582293271726 --zip skill/m365-command-planner.zip
python3 skill/create_skill.py --api-mode widget --live --zip skill/m365-command-planner.zip
```

## Current dev app replace runbook

The current dev app is:

```bash
export GE_PROJECT=saib-ai-playground
export GE_PROJECT_NUMBER=288406675721
export GE_LOCATION=global
export GE_ENGINE=ge-msft-plugin-test_1782382759735
export GE_WIDGET_CONFIG_ID=cd8248bf-0b65-487d-9a81-fdd48f3912e7
export GE_WIDGET_SERVER_TOKEN=CAMSAh0H
```

Optional private skill id hints. Stable labels are valid; live widget replace resolves the current
numeric ids before deleting:

```bash
export GE_COMMAND_PLANNER_AGENT_ID=m365-command-planner
export GE_SURFACE_COMMANDER_AGENT_ID=m365-surface-commander
```

Rebuild and validate the local bundles before any live mutation:

```bash
cd /home/user/ge-msft
bun run skills:check
python3 skill/validate_skill_bundles.py

cd /home/user/ge-msft/skill
./build_zip.sh m365-command-planner
./build_zip.sh m365-surface-commander
cd /home/user/ge-msft
```

Provide a short-lived widget bearer token using either a DevTools-exported widget request:

```bash
python3 skill/extract_widget_credentials.py /path/to/widget-request.curl \
  --env-file /tmp/ge-widget.env \
  --project "$GE_PROJECT" \
  --agent-planner "$GE_COMMAND_PLANNER_AGENT_ID" \
  --agent-surface "$GE_SURFACE_COMMANDER_AGENT_ID"

source /tmp/ge-widget.env
```

The guided updater accepts the same request as a file, stdin, one-line cURL, or a pasted multi-line
cURL/HAR block:

```bash
# Prompt for a path, one-line cURL, stdin, or PASTE mode.
scripts/update-ge-widget-skills.sh

# Force paste mode; finish the pasted block with __GE_WIDGET_CURL_END__ on its own line.
scripts/update-ge-widget-skills.sh --paste-curl
mise run ge:skills:update:paste

# Use a saved DevTools "Copy as cURL" text file or HAR export.
scripts/update-ge-widget-skills.sh --credentials-file /tmp/ge-widget-request.curl

# Pipe a copied request directly.
pbpaste | scripts/update-ge-widget-skills.sh --credentials-file -
```

The updater reuses `GE_WIDGET_BEARER_TOKEN_FILE` or `/tmp/ge-widget-token` when the captured widget
JWT is still valid and has at least 120 seconds remaining. That avoids repeated DevTools copy/paste
during a short working session. Force a new widget request when switching accounts, switching Gemini
Enterprise apps, or repairing a failed token:

```bash
scripts/update-ge-widget-skills.sh --force-token-refresh
mise run ge:skills:update:force-token
```

Tune the safety window with `--min-token-ttl <seconds>` or `GE_WIDGET_MIN_TTL_SECONDS`. The widget
bearer itself is short-lived by design; the script can reuse a valid token but cannot extend its
server-side expiry.

Or, when the documented widget-token probe succeeds for the signed-in WIF account:

```bash
CLOUDSDK_CONFIG=/home/user/ge-msft/.gcloud \
gcloud auth login \
  --login-config=/home/user/ge-msft/.gcloud/saib-wif-login-config.json \
  --no-launch-browser

python3 skill/acquire_widget_token.py \
  --live \
  --host both \
  --use-gcloud \
  --config-id "$GE_WIDGET_CONFIG_ID" \
  --project-number "$GE_PROJECT_NUMBER" \
  --connector-id msft-onedrive-fed_1779469629030 \
  --location "$GE_LOCATION" \
  --quota-project "$GE_PROJECT" \
  --write-widget-token-file /tmp/ge-widget-token

export GE_WIDGET_BEARER_TOKEN_FILE=/tmp/ge-widget-token
```

Then list, replace, and list again:

```bash
python3 skill/update_skills.py --api-mode widget --live --list
python3 skill/update_skills.py --api-mode widget --replace --yes --live
python3 skill/update_skills.py --api-mode widget --live --list
```

For widget `--replace`, `skill/update_skills.py` treats configured numeric agent ids as hints. It
first lists visible private skill-agent views and resolves current delete targets by the stable
bundle names `m365-command-planner` and `m365-surface-commander`. If a prior failed replace left a
new shell agent behind, the same name-based pass deletes that duplicate before creating one clean
replacement.

`--replace --yes --live` is destructive for the selected numeric skill agents. It should print the
new `VITE_GE_COMMAND_PLANNER_SKILL` and `VITE_GE_SURFACE_COMMANDER_SKILL` values. The guided
`scripts/update-ge-widget-skills.sh` flow writes those non-secret values, plus the widget config id
and server token, into `packages/web-shell/.env` automatically after a successful replacement. The
short-lived widget bearer token remains only in `/tmp` and is never copied into the web-shell env.

The same successful replacement also writes non-secret bundle provenance:

| Env key | Meaning |
| --- | --- |
| `VITE_GE_COMMAND_PLANNER_SKILL_VERSION` / `VITE_GE_SURFACE_COMMANDER_SKILL_VERSION` | Version parsed from the local `SKILL.md` frontmatter. |
| `*_SKILL_SOURCE_SHA256` | Hash of the rebuilt source zip before upload stamping. This is the drift check: if the local zip hash changes, regenerate and replace the GE skill. |
| `*_SKILL_SHA256` | Hash of the temporary stamped zip that was uploaded. This differs from the source hash because the upload zip includes the numeric agent id in `SKILL.md`. |
| `VITE_GE_SKILL_SOURCE_BUNDLE_SET_SHA256` | Combined source hash across the two bundled skills. |
| `VITE_GE_SKILL_UPLOAD_BUNDLE_SET_SHA256` | Combined uploaded-stamped hash across the two bundled skills. |

At upload time, `skill/update_skills.py` creates a temporary zip and injects computed frontmatter
into its `SKILL.md`:

```yaml
x-ge-msft-upload:
  buildId: 'm365-command-planner@17644156643933695033+<source-sha12>'
  agentId: '17644156643933695033'
  resource: 'projects/.../assistants/default_assistant/agents/17644156643933695033'
  sourceZipSha256: '<sha256 of the rebuilt source zip>'
  sourceVersion: '1.1'
```

This identifier is computed only after the widget create/list step has resolved the actual numeric
agent id. The checked-in source skill folders and source zips are left untouched.

To sync `.env` from the current visible skill list without replacing anything:

```bash
mise run ge:skills:list
# or
scripts/update-ge-widget-skills.sh --list-only
```

Use `--no-write-web-shell-env` to opt out, or `--web-shell-env <path>` to update a different env
file. List-only sync intentionally does not update hash/version keys; it can verify the current
agent ids, but not the exact bundle bytes currently stored inside those private agents.

## WidgetAcquireAccessToken probe

The official RPC reference also documents deprecated/internal
`WidgetService.WidgetAcquireAccessToken`. Its schema does **not** directly say "mint widget bearer
token"; it says the method proxies `DataConnectorService.AcquireAccessToken` and returns:

- `acquireAccessTokenResponse.accessToken`: the per-user connector access token.
- `uToken`: an additional token from the widget response.

Because that response can include connector credentials, the repo treats this as a redacted probe,
not as default credential plumbing. The probe writes a token file only when `uToken` validates as the
same widget JWT shape used by the private skill endpoints:

```bash
python3 skill/acquire_widget_token.py \
  --live \
  --host both \
  --use-gcloud \
  --config-id "$GE_WIDGET_CONFIG_ID" \
  --project-number "$GE_PROJECT_NUMBER" \
  --connector-id msft-onedrive-fed_1779469629030 \
  --location global \
  --quota-project saib-ai-playground \
  --write-widget-token-file /tmp/ge-widget-token
```

If `gcloud auth print-access-token` reports `invalid_grant: Refresh token has expired`, re-login
without `--update-adc`:

```bash
CLOUDSDK_CONFIG=/home/user/ge-msft/.gcloud \
gcloud auth login \
  --login-config=/home/user/ge-msft/.gcloud/saib-wif-login-config.json \
  --no-launch-browser
```

Then rerun the probe. The probe prints HTTP status, response keys, token issuer/audience/expiry
metadata, and whether `uToken` is a valid widget bearer. It never prints token values.

If the API reports that no quota project is set, pass `--quota-project <project-id>`. For the dev
tenant used here, that is `saib-ai-playground`; the probe sends it as `x-goog-user-project`.

Live result for the current dev app:

- The documented full connector resource name was rejected with `INVALID_ARGUMENT`.
- The accepted connector name shape was the widget-erased resource:

```text
collections/msft-onedrive-fed_1779469629030/dataConnector
```

- The accepted payload shape was the direct top-level RPC request:

```json
{
  "configId": "cd8248bf-0b65-487d-9a81-fdd48f3912e7",
  "location": "locations/global",
  "additionalParams": {
    "token": "-",
    "origin": "ORIGIN_UNSPECIFIED"
  },
  "acquireAccessTokenRequest": {
    "name": "collections/msft-onedrive-fed_1779469629030/dataConnector"
  }
}
```

- Both `content-discoveryengine.googleapis.com` and `discoveryengine.googleapis.com` accepted that
  shape when the request included `x-goog-user-project: saib-ai-playground`.
- The response contained `acquireAccessTokenResponse.accessToken`, which decoded as a Microsoft
  Graph connector token (`aud=https://graph.microsoft.com`) for tenant
  `99ae41c3-878e-4b22-9467-b8126cb10218`.
- The response did **not** include `uToken`, so it did not produce the widget bearer required for
  `widgetCreateAgent`, `widgetDeleteAgent`, `widgetListAvailableAgentViews`, or skill zip upload.

Conclusion: `WidgetAcquireAccessToken` is useful evidence for connector-token acquisition, but it
does **not** replace the current explicit widget-bearer extraction step for private skill lifecycle
operations.

The updater prints add-in env values after upload:

```text
VITE_GE_COMMAND_PLANNER_SKILL=m365-command-planner=projects/.../assistants/default_assistant/agents/17573173582293271726
VITE_GE_SURFACE_COMMANDER_SKILL=m365-surface-commander=projects/.../assistants/default_assistant/agents/7404511736383961129
```

Those values are the source of truth for send-time routing.

## Send-time reference contract

The add-in must use both parts:

1. `skillsSpec.skills[].name` uses the full skill resource name:

```json
{
  "skillsSpec": {
    "skills": [
      {
        "name": "projects/.../locations/global/collections/default_collection/engines/ge-msft-plugin-test_1782382759735/assistants/default_assistant/agents/17573173582293271726"
      }
    ]
  }
}
```

2. The query text is prefixed with the widget mention marker:

```text
[m365-command-planner](mention://?uri=17573173582293271726)
```

The full resource name mounts the skill. The mention marker matches how the Gemini Enterprise UI
routes visible skills and materially improves deterministic skill activation. The runtime tests pin
that planner turns only mount/mention `m365-command-planner`, and command-loop turns only
mount/mention `m365-surface-commander`.

## Important distinction: public `agentsSpec` is not this skill mount

The current public RPC reference includes `StreamAssistRequest.agents_spec`. The live Gemini
Enterprise skill flow we use still requires the widget-style `skillsSpec.skills[].name` plus the
mention marker above. Do not replace `skillsSpec` with `agentsSpec` until a live test proves that
the same uploaded `skillAgentDefinition` agent is invoked with equivalent behavior.

## Safety

- Widget tokens are short-lived local developer credentials. Store only in a temp file with mode
  `0600`; never commit.
- Do not automate browser cookies or XSRF flows.
- Do not treat widget permissions as tenant-admin IAM. Use public/admin API mode for CI when the
  tenant grants the appropriate Google IAM role.
- Destructive operations require `--yes`.
