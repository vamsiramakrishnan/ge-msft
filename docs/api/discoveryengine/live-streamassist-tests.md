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
- `m365-surface-commander` can be selected with `agentsSpec` and an explicit mention and returns `cmd` fences for strict
  Excel, Word, PowerPoint, Outlook, and injection-resistance command-loop fixtures;
- `m365-command-planner` can be selected with `agentsSpec` and an explicit mention and returns `plan` fences for
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

## Session mode and workflow comparison

Compare the same private commander routing and command-generation fixtures with and without a
provider conversation:

```bash
bun run test:streamassist:modes
```

`GE_LIVE_STREAMASSIST_SESSION_MODES=conversation,sessionless` expands each selected scenario into
both request modes. Conversation samples start fresh; sessionless samples send `isSessionLess: true`
and omit `session`. The harness reads the wire response before any production-client filtering. It
fails a sessionless sample that returns `sessionInfo.session`, and a conversation sample that does
not return one. This checks the API response contract. Confirming absence from saved UI history
still requires a tenant check; an absent response field is not proof of storage behavior.

The `commander-verified-program` scenario requests a bound capture/reconcile/materialize program.
It checks the complete program with the production grammar parser and requires the configured
commander identity in `invokedSkills`. A generic skill-name match cannot pass this routing probe.
The evidence includes time to the first complete parseable program. This is a streaming observation;
execution still waits for the whole response so later output cannot change program validity.

Reports group repetitions by scenario and session mode. Each timing stage includes its sample
count and nearest-rank p50/p95. Missing first-token or parseable-program timestamps remain missing,
so a failed response cannot lower latency by contributing a synthetic zero.

To measure through actual command execution and verification:

```bash
# Supply an existing signed-in user's WIF token through the environment or a token file.
# This command never runs a login flow or falls back to a metadata/service-account identity.
bun run test:command-workflows:live
```

This opt-in suite runs the production `StreamAssistClient`, `AssistSession`, recovery checkpoints,
approval boundary and DuckDB WASM against a synthetic workbook. It compares five variants of the
same exact-decimal reconciliation:

| Provider mode | Context | Command encoding |
| --- | --- | --- |
| Conversation | Transcript | Sequential artifact handoffs |
| Sessionless | Transcript | Sequential artifact handoffs |
| Sessionless | Current execution state | Sequential artifact handoffs |
| Conversation | Transcript | One bound program |
| Sessionless | Current execution state | One bound program |

Each sample gets fresh runtime state and workbook contents. The fixture includes an amount above
JavaScript's safe-integer range, decimal aggregation, a variance and mismatched currencies. Results
must preserve those exact values and classifications. Private skill invocation, request mode,
returned session metadata, approvals, effects, model calls, correction turns and task completion are recorded. Encoding fidelity is checked separately: the one-program variant must finish in one model call, and the requested sequential variant in four. A successful task that ignores its assigned variant cannot silently pass the comparison.

Required inputs are `GE_ENGINE`, `GE_SURFACE_COMMANDER_SKILL`, `GE_PROJECT_NUMBER` (or a numeric
project in the skill resource), and `GE_WIF_ACCESS_TOKEN`/`GE_ACCESS_TOKEN` or
`GE_WIF_ACCESS_TOKEN_FILE`/`GE_ACCESS_TOKEN_FILE`. `GE_LOCATION` defaults to `global`.
`GE_USER_PROJECT`/`GE_PROJECT` sets the quota header. `GE_LIVE_COMMAND_REPETITIONS` accepts 1–10
samples per variant and defaults to 3. This workflow suite reads explicitly exported variables;
it does not load browser/widget credentials or configuration files automatically.

The report is `dist/probes/command-workflows-live.json`. It separates first-token and parseable
program latency, cumulative model duration, deterministic DuckDB calculation, approval wait,
host application plus readback verification, and verified completion. Per-stage p50/p95 include
their sample counts. Success counts accompany timing; compare them before optimizing latency.
Query bytes describe submitted query text, not billed provider tokens or provider-side history.
No prompts, raw replies, source rows or bearer tokens are written to this report.

**The provider and DuckDB are live; Office and the approval responder are simulated.** The approval
wait measurement is the fixture callback overhead, not human decision time. Host timing therefore
does not measure Office.js or a real coauthor. Tenant saved-history behavior, real host latency and
human review remain separate validation gates. Neither live suite runs during the default tests.
