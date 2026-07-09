# Live StreamAssist integration tests

These tests are opt-in. Normal `bun run test` does not call Gemini Enterprise.

Run them only when you want to cross-validate the current add-in skill wiring against the real
Gemini Enterprise `widgetStreamAssist` endpoint.

## Inputs

Required:

- `GE_WIDGET_BEARER_TOKEN_FILE` or `GE_WIDGET_BEARER_TOKEN`
- `GE_WIDGET_CONFIG_ID`
- `GE_WIDGET_SERVER_TOKEN`
- `GE_ENGINE`
- `GE_SURFACE_COMMANDER_SKILL`
- `GE_COMMAND_PLANNER_SKILL`

Optional:

- `GE_LOCATION` (defaults to `global`)
- `GE_TIME_ZONE` (defaults to `UTC`)
- `GE_LIVE_STREAMASSIST_SCENARIOS` comma-separated scenario ids
- `GE_LIVE_STREAMASSIST_REQUIRE_CODE=1` to fail when the code-execution probe does not emit
  executable-code/code-result parts

The test also reads these fallback files when present:

- `/tmp/ge-widget.env`
- `packages/web-shell/.env`

`packages/web-shell/.env` supplies the current dev engine and skill resource names. `/tmp/ge-widget.env`
should contain only short-lived local widget credentials produced from a browser-authenticated Gemini
Enterprise session. Do not commit either bearer tokens or generated evidence that contains sensitive
live prompts.

## Command

```bash
bun run test:streamassist:live
```

If the local widget token is missing or expired, use the interactive login preflight:

```bash
bun run test:streamassist:live:login
```

The preflight opens or prints the Gemini Enterprise URL, asks you to paste one authenticated
`content-discoveryengine.googleapis.com` DevTools cURL/HAR request, extracts only the short-lived
widget bearer/config into `/tmp/ge-widget.env`, then launches the same Vitest suite. It deliberately
does not automate browser cookies, XSRF state, or Google session secrets.

Run one or a few scenarios while debugging:

```bash
GE_LIVE_STREAMASSIST_SCENARIOS=smoke-basic,commander-excel-visualize bun run test:streamassist:live
```

The login preflight supports the same scenario filter:

```bash
bun run test:streamassist:live:login -- --scenarios smoke-basic,commander-excel-visualize
```

The suite verifies:

- the widget StreamAssist endpoint is reachable and returns a streamed answer;
- hosted code-execution behavior is observable on analytical prompts;
- `m365-surface-commander` can be mounted with `skillsSpec` and returns `cmd` fences for strict
  Excel, Word, PowerPoint, Outlook, and injection-resistance command-loop fixtures;
- `m365-command-planner` can be mounted with `skillsSpec` and returns `plan` fences for
  single-surface and cross-surface handoff requests;
- both planner and commander can be mounted in the same session without breaking the response
  schema.

Current scenario ids:

```text
smoke-basic
smoke-code-exec-observe
commander-excel-visualize
commander-word-surgical
commander-powerpoint-shape
commander-outlook-draft-only
commander-injection-resistance
planner-single-surface
planner-cross-surface
multi-skill-mount
```

If the bearer token has expired, refresh `/tmp/ge-widget.env` and rerun the command.

## Evidence

Live runs write:

```text
dist/probes/streamassist-live.json
```

The evidence file contains scenario status, failure messages, response hashes, small synthetic
previews, invoked skill names, session names, chunk counts, and whether code execution parts were
observed. It does not store bearer tokens.
