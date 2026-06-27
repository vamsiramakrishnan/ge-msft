# Example session: progressive full-file context

Use this pattern when the visible snapshot or selected range is too small for the task, especially
for workbook-wide reconciliation, file-level validation, or analysis that may benefit from
StreamAssist-hosted file handling.

## User task

> Analyze the whole workbook, find schedule risks, and create an appropriate chart.

## First turn

The model does not guess from the snapshot. It asks the host for a context strategy:

```cmd
context analytical full-scope upload-preferred code-execution-preferred
```

## Result shape

The add-in returns a `result` entry with:

- `strategy.scope = whole-artifact`
- `strategy.transfer = upload-candidate`
- `strategy.analysis = code-execution-candidate`
- accepted file extensions and MIME types
- max byte limits
- guardrails that uploads are explicit host/user actions

## Next turn

If the host/user attaches the workbook and supplies a structured file id, use that host-provided
grounding in the next StreamAssist turn. Do not invent file ids and do not emit upload or code
execution commands. If no file id is supplied, continue with bounded `outline`, `read`, and `search`
commands.
