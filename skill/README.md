# Gemini Enterprise skill tooling — create & test

Tooling to **create** a Gemini Enterprise custom skill programmatically and **test/refine** it
against the `streamAssist` API in isolation (no other data sources connected).

````
ge-skill-tooling/
├── create_skill.py            # create/upload a skill via public API or widget API mode
├── extract_widget_credentials.py # dev-only helper: extract widget env from a saved cURL/HAR
├── test_skill.py              # multi-surface live test harness (+ offline self-check)
├── de_stub.py                 # streamAssist response stub + robust reader (thoughts/citations/…)
├── fixtures.py                # mock M365 docs: Excel analysis, Outlook thread, Word contract
├── build_zip.sh               # (re)build a skill bundle zip: ./build_zip.sh <skill-dir>
├── requirements.txt
├── m365-surface-commander/    # EXECUTOR bundle — emits the ```cmd algebra (agentskills.io format)
│   ├── SKILL.md
│   └── references/  scripts/  assets/
└── m365-command-planner/      # PLANNER bundle — free text -> a confirmable ```plan block
    ├── SKILL.md
    └── references/  scripts/
````

## Runtime skills and developer skill

The command surface (`/` verbs + `@` mentions in the add-in) is carried into Gemini Enterprise as
**two skills, mounted per-turn via `skillsSpec`**:

- **`m365-command-planner`** — the **front door**. Turns a free-text `/verb @mentions …` request
  into a structured, parseable ` ```plan ` block (intent · scope · ordered steps · exclusions ·
  grounding), which the add-in renders for a one-tap confirm. It never touches the document.
- **`m365-surface-commander`** — the **executor**. Takes the confirmed plan + a live document
  snapshot and emits the ` ```cmd ` command algebra the add-in applies as reviewable changes.

Route by complexity: a simple `verb + mentions` request goes straight to the executor; free text
with constraints/exclusions goes through the planner first. Build either bundle with
`./build_zip.sh <skill-dir>` and create it with `create_skill.py --zip <skill-dir>.zip`.

This repo also includes **`m365-release-operator`**, a developer/operator skill for the local coding
harness. It guides readiness checks, Bun bootstrap flows, Entra redirect sync, sideload/catalog
deployment, widget cURL/HAR paste handling, and skill provisioning. It is **not** mounted by the
Office add-in at runtime and is not part of the normal widget skill replacement flow. Build it
explicitly when needed:

```bash
cd /home/user/ge-msft/skill
./build_zip.sh m365-release-operator
```

See `../docs/api/discoveryengine/skills-and-agents.md` for the verified create/mount lifecycle and
`../docs/api/discoveryengine/widget-service-skills.md` for the widget API contract, including the
send-time `skillsSpec` + `mention://` reference format. See also
the **skill ↔ workspace parity** tasks (keep `parse_commands.py` / `parse_plan.py` in lockstep with
`packages/contracts` + `packages/runtime`; the TypeScript side is authoritative).

## Prerequisites

```bash
pip install -r requirements.txt
gcloud auth application-default login        # public API mode; needs Discovery Engine agent IAM
```

Configure the target engine (defaults point at the dev engine; override for your project):

```bash
export GE_PROJECT=your-project-id
export GE_PROJECT_NUMBER=123456789012
export GE_LOCATION=global
export GE_ENGINE=your-engine_1700000000000
```

## Create / upload a skill

There are two API modes:

- `--api-mode public` uses `discoveryengine.googleapis.com` with OAuth/ADC. It is useful only when
  the public assistant-agent API is enabled for the target resource. It does **not** update the
  numeric UI-created Gemini Enterprise widget skills used by the current dev app.
- `--api-mode widget` mirrors the Gemini Enterprise web UI (`widgetCreateAgent`,
  `widgetListAvailableAgentViews`, resumable `files:upload`). It uses a short-lived Vertex AI
  Search widget JWT plus `GE_WIDGET_CONFIG_ID`; keep it dev-only and never commit the token.

```bash
# (re)build the bundle zip from the m365-surface-commander/ directory
./build_zip.sh

# Public API auth/permission probe; no mutation.
python3 create_skill.py --api-mode public --live --list

# Method B (default): create a shell agent, then raw-upload the zip (server unpacks SKILL.md +
# references/scripts/assets into instruction + subfiles)
python3 create_skill.py --api-mode public --agent-id m365-surface-commander --zip m365-surface-commander.zip

# Method A: single-file skill — push one markdown file as the instruction (no bundle)
python3 create_skill.py --api-mode public --single-file m365-surface-commander/SKILL.md

# Widget/dev update of existing UI-created skills. First export one authenticated
# content-discoveryengine.googleapis.com request from DevTools as cURL or HAR, then extract only the
# short-lived widget JWT/config into local env exports:
python3 extract_widget_credentials.py /path/to/widget-request.curl \
  --env-file /tmp/ge-widget.env \
  --project saib-ai-playground \
  --agent-surface 8870098647237058037 \
  --agent-planner 17573173582293271726

# Run the printed export lines, then replace the two skills with the current zip bundles:
source /tmp/ge-widget.env
python3 update_skills.py --api-mode widget --replace --yes --live

# List visible private widget skills and the exact resource/mention references.
python3 update_skills.py --api-mode widget --live --list

# Probe the documented WidgetAcquireAccessToken path. This is redacted by default and writes only a
# returned uToken that validates as a widget bearer JWT.
python3 acquire_widget_token.py \
  --live \
  --host both \
  --use-gcloud \
  --config-id "$GE_WIDGET_CONFIG_ID" \
  --project-number "$GE_PROJECT_NUMBER" \
  --connector-id msft-onedrive-fed_1779469629030 \
  --location global \
  --quota-project saib-ai-playground \
  --write-widget-token-file /tmp/ge-widget-token

# Useful flags
python3 create_skill.py --api-mode public --replace      # delete an existing agent first
python3 create_skill.py --api-mode public --share        # set sharingConfig.scope=ALL_USERS
```

### Dev app runbook: delete existing skills and upload latest bundles

Use this for the current `saib-ai-playground` dev Gemini Enterprise app. It deletes the two selected
private widget skills, creates replacements, uploads the freshly built zip bundles, verifies them,
and writes the new `VITE_GE_*_SKILL` values into `packages/web-shell/.env` by default.

The guided script wraps the same steps, opens the Gemini Enterprise app URL, asks for a saved
DevTools cURL/HAR request, extracts the short-lived widget token into `/tmp`, and then runs the
list/replace flow:

```bash
cd /home/user/ge-msft
scripts/update-ge-widget-skills.sh

# With mise:
mise run ge:skills:update
```

Use `scripts/update-ge-widget-skills.sh --dry-run` or `mise run ge:skills:dry-run` to verify the
target and zips without live mutation. Use `--list-only` or `mise run ge:skills:list` to list
visible private widget skills and sync the current refs into `packages/web-shell/.env` without
deleting or recreating anything.

The guided updater reuses `GE_WIDGET_BEARER_TOKEN_FILE` or `/tmp/ge-widget-token` when the captured
widget JWT is still valid and has enough lifetime left. It only opens Gemini Enterprise and asks for
a new DevTools cURL/HAR when the token is missing, invalid, near expiry, or when you pass:

```bash
scripts/update-ge-widget-skills.sh --force-token-refresh
mise run ge:skills:update:force-token
```

Use `--min-token-ttl <seconds>` or `GE_WIDGET_MIN_TTL_SECONDS` to adjust the reuse threshold. This
does not extend the widget token's server-side lifetime; it simply avoids unnecessary re-auth prompts
while the current token remains usable.

Successful uploads also stamp the uploaded `SKILL.md` frontmatter with computed provenance:

```yaml
x-ge-msft-upload:
  buildId: 'm365-surface-commander@15135346478580045234+<source-sha12>'
  agentId: '15135346478580045234'
  resource: 'projects/.../assistants/default_assistant/agents/15135346478580045234'
  sourceZipSha256: '<sha256 of the committed source zip>'
  sourceVersion: '1.1'
```

The source files under `skill/m365-*` are not mutated. The updater creates a temporary stamped zip
after Gemini Enterprise returns the numeric agent id, uploads that zip, and writes non-secret
provenance into the web-shell env:

- `*_SKILL_VERSION`: version found in the source `SKILL.md` frontmatter.
- `*_SKILL_SOURCE_SHA256`: hash of the rebuilt source zip; use this to know whether a replacement
  upload is needed.
- `*_SKILL_SHA256`: hash of the temporary stamped zip that was actually uploaded.
- `VITE_GE_SKILL_SOURCE_BUNDLE_SET_SHA256` and `VITE_GE_SKILL_UPLOAD_BUNDLE_SET_SHA256`: combined
  hashes across the two bundled skills.

`--list-only` deliberately syncs only the current skill references and widget config. It does not
rewrite hash/version values because listing can prove the current numeric agent ids, but it cannot
prove which local bundle was uploaded into those agents.

Pass `--no-write-web-shell-env` if you want to print the values without updating
`packages/web-shell/.env`. Pass `--web-shell-env <path>` to update a different env file.

During widget `--replace`, configured numeric `GE_*_AGENT_ID` values are only hints. The updater
lists visible private skills, resolves the current delete targets by the stable names
`m365-command-planner` and `m365-surface-commander`, deletes all matching candidates, then creates
one clean replacement for each bundle. This prevents stale-id failures after prior replace runs.

The guided script accepts the widget request in four forms:

```bash
# Prompt for a path, one-line cURL, or PASTE mode.
scripts/update-ge-widget-skills.sh

# Force paste mode for a multi-line cURL/HAR block; finish with __GE_WIDGET_CURL_END__.
scripts/update-ge-widget-skills.sh --paste-curl
mise run ge:skills:update:paste

# Use a saved DevTools cURL/HAR file.
scripts/update-ge-widget-skills.sh --credentials-file /tmp/ge-widget-request.curl

# Pipe a copied request from stdin.
pbpaste | scripts/update-ge-widget-skills.sh --credentials-file -
```

1. Rebuild and validate locally.

```bash
cd /home/user/ge-msft
bun run skills:check
python3 skill/validate_skill_bundles.py

cd /home/user/ge-msft/skill
./build_zip.sh m365-surface-commander
./build_zip.sh m365-command-planner
```

2. Export the app and skill ids. If the ids are stale, run the list command in step 4 first and
   replace them with the current numeric `name` values.

```bash
cd /home/user/ge-msft
export GE_PROJECT=saib-ai-playground
export GE_PROJECT_NUMBER=288406675721
export GE_LOCATION=global
export GE_ENGINE=ge-msft-plugin-test_1782382759735
export GE_WIDGET_CONFIG_ID=cd8248bf-0b65-487d-9a81-fdd48f3912e7
export GE_WIDGET_SERVER_TOKEN=CAMSAh0H

export GE_COMMAND_PLANNER_AGENT_ID=m365-command-planner
export GE_SURFACE_COMMANDER_AGENT_ID=m365-surface-commander
```

3. Provide a widget bearer token. Preferred dev path: export one authenticated
   `content-discoveryengine.googleapis.com` request from Gemini Enterprise DevTools as cURL/HAR and
   extract it.

```bash
python3 skill/extract_widget_credentials.py /path/to/widget-request.curl \
  --env-file /tmp/ge-widget.env \
  --project "$GE_PROJECT" \
  --agent-planner "$GE_COMMAND_PLANNER_AGENT_ID" \
  --agent-surface "$GE_SURFACE_COMMANDER_AGENT_ID"

source /tmp/ge-widget.env
```

If the documented widget token probe works in your session, this is also acceptable:

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

4. Verify the target and run the destructive replace.

```bash
python3 skill/update_skills.py --api-mode widget --live --list
python3 skill/update_skills.py --api-mode widget --replace --yes --live
python3 skill/update_skills.py --api-mode widget --live --list
```

5. If you are not using the guided script, write the printed replacement skill references into
   `packages/web-shell/.env`:

```text
VITE_GE_COMMAND_PLANNER_SKILL=...
VITE_GE_SURFACE_COMMANDER_SKILL=...
```

The guided script does this automatically after a successful live replace.

The public API and widget API are deliberately separate. The current Gemini Enterprise dev app uses
UI-created widget skills, so the operational update path is widget mode. Do not automate browser
cookies, SAPISIDHASH, or XSRF state. `extract_widget_credentials.py` accepts only a saved request
you explicitly export from DevTools, validates that the Authorization header is a Vertex AI Search
widget JWT, stores that short-lived token in a local `0600` temp file, and prints the env vars needed
by `update_skills.py`.

## Test / refine a skill

The harness drives `streamAssist` with **only the skill connected and an empty `toolsSpec`** (no web
grounding, no data stores), then simulates the Office add-in's multi-turn loop against a mock
document — applying the model's commands and feeding back a `result` block each turn.

```bash
# live, against the deployed skill, per surface
python3 test_skill.py --agent m365-surface-commander --surface excel
python3 test_skill.py --agent m365-surface-commander --surface email
python3 test_skill.py --agent m365-surface-commander --surface contract --raw

# offline harness self-check (no API) — proves reader+parser+fixtures are sound
python3 test_skill.py --stub
```

Each run prints per-turn `[CMD]`/`[PROSE]`, the parsed commands, applied effects, and a metrics
block: `cmd_blocks`, `errors`, `prose_only`, `grounding_leak`, `done`.
**`grounding_leak: false` confirms isolation** — no data source contributed to the answer.

## What we learned (refinement notes)

- **Isolation:** empty `toolsSpec` reliably isolates the skill (no grounding/citations leak).
- **Highest-leverage reliability lever is host-side:** inject each available verb's _exact usage_
  (not just verb names) in the per-turn `<capabilities>` block. This eliminated verb/syntax drift
  (`reply(body=…)` → `mail "…"`, malformed `suggest`, etc.). See `render_caps()` in `test_skill.py`.
- **Residual:** as a `skillsSpec`-layered skill the model tends to answer the first turn in prose,
  then emits correct commands after a re-prompt. The add-in should: re-prompt on a no-command turn,
  and **not** honor `done` if the same block had parse errors.
- The stub (`de_stub.py`) reproduces the real wire complications — token-streamed text, thoughts,
  `textGroundingMetadata` citations, `inlineData` suggestions, and split code fences — so the reader
  is exercised against them offline.
