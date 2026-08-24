# Live StreamAssist integration tests

These tests are opt-in. Normal `bun run test` does not call Gemini Enterprise.

Run them only when you want to cross-validate the current add-in skill wiring against the real
Gemini Enterprise public `streamAssist` endpoint. This is the production-auth path: the suite uses
the signed-in Workforce Identity Federation principal's Google access token. Widget authentication
is retained only as an explicit compatibility mode for private skill catalog diagnostics.

## Inputs

Required for the default WIF transport:

- `GE_ENGINE`
- `GE_SURFACE_COMMANDER_SKILL`
- `GE_COMMAND_PLANNER_SKILL`

The access token resolves in this order:

1. `GE_WIF_ACCESS_TOKEN` or `GE_ACCESS_TOKEN`;
2. `GE_WIF_ACCESS_TOKEN_FILE` or `GE_ACCESS_TOKEN_FILE`;
3. `gcloud auth print-access-token`, using `GE_CLOUDSDK_CONFIG`, `CLOUDSDK_CONFIG`, or the repo's
   `.gcloud` directory.

For explicit widget compatibility mode (`bun run test:streamassist:widget`), also supply
`GE_WIDGET_BEARER_TOKEN_FILE`/`GE_WIDGET_BEARER_TOKEN`, `GE_WIDGET_CONFIG_ID`, and
`GE_WIDGET_SERVER_TOKEN`.

Optional:

- `GE_PROJECT_NUMBER` (otherwise derived from the configured skill resource names)
- `GE_PROJECT` or `GE_USER_PROJECT` for the quota header (otherwise read from gcloud)
- `GE_LOCATION` (defaults to `global`)
- `GE_TIME_ZONE` (defaults to `UTC`)
- `GE_LIVE_STREAMASSIST_SCENARIOS` comma-separated scenario ids
- `GE_LIVE_STREAMASSIST_REPETITIONS` runs every selected scenario 1–10 times (defaults to `1`)
- `GE_LIVE_STREAMASSIST_REQUIRE_CODE=1` to fail when the code-execution probe does not emit
  executable-code/code-result parts
- `GE_LIVE_STREAMASSIST_REQUIRE_CLOSED_FENCE=1` to make a recoverable unclosed `cmd` or `plan`
  fence fail the live test instead of recording a protocol warning

The test also reads these fallback files when present:

- `/tmp/ge-widget.env`
- `packages/web-shell/.env`

`packages/web-shell/.env` supplies the current dev engine and skill resource names. `/tmp/ge-widget.env`
is read only for explicit widget compatibility runs and skill catalog uploads. Do not commit either
bearer tokens or generated evidence that contains sensitive live prompts.

## Command

```bash
bun run test:streamassist:live
```

Run the focused first-token benchmark (four representative scenarios, three samples each):

```bash
bun run test:streamassist:latency
```

Widget-only compatibility check:

```bash
bun run test:streamassist:widget
```

The separate skill uploader may ask for a browser-authenticated widget cURL/HAR because private
skill catalog list/upload is not the public `streamAssist` data path. Never substitute that
five-minute widget JWT for the WIF access token used by the add-in and default live suite.

Run one or a few scenarios while debugging:

```bash
GE_LIVE_STREAMASSIST_SCENARIOS=smoke-basic,commander-excel-visualize bun run test:streamassist:live
```

The suite verifies:

- the public WIF-authenticated StreamAssist endpoint is reachable and returns a streamed answer;
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

If gcloud cannot mint a token, refresh the repo WIF login with `bun run setup:gcloud:wif`.

## Evidence

Live runs write:

```text
dist/probes/streamassist-live.json
```

The evidence file contains scenario status, failure messages, response hashes, small synthetic
previews, invoked skill names, session names, chunk counts, whether code execution parts were
observed, and wire-level latency for response headers, first parsed chunk, first non-thought answer
token, first visible text (including the activity text the client emits from thought frames), and
completion. Its summary records median and p95 visible-text, answer-token, and total duration.
An opening `cmd` or `plan` fence that reaches end-of-response is recorded in `protocolWarnings`:
the production parsers deliberately recover only this bounded whole-response shape. The report does
not store bearer tokens.
