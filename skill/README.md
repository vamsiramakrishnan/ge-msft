# Gemini Enterprise skill tooling — create & test

Tooling to **create** a Gemini Enterprise custom skill programmatically and **test/refine** it
against the `streamAssist` API in isolation (no other data sources connected). Built and verified
end-to-end against a live GE engine.

```
ge-skill-tooling/
├── create_skill.py            # create/upload a skill via the authenticated GE API
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
```

## The two skills

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

See `../docs/api/discoveryengine/skills-and-agents.md` for the verified create/mount lifecycle and
the **skill ↔ workspace parity** tasks (keep `parse_commands.py` / `parse_plan.py` in lockstep with
`packages/contracts` + `packages/runtime`; the TypeScript side is authoritative).

## Prerequisites

```bash
pip install -r requirements.txt
gcloud auth application-default login        # ADC; identity needs agents create/update on the engine
```

Configure the target engine (defaults point at the dev engine; override for your project):

```bash
export GE_PROJECT=your-project-id
export GE_PROJECT_NUMBER=123456789012
export GE_LOCATION=global
export GE_ENGINE=your-engine_1700000000000
```

## Create / upload a skill

```bash
# (re)build the bundle zip from the m365-surface-commander/ directory
./build_zip.sh

# Method B (default): create a shell agent, then raw-upload the zip (server unpacks SKILL.md +
# references/scripts/assets into instruction + subfiles)
python3 create_skill.py --agent-id m365-surface-commander --zip m365-surface-commander.zip

# Method A: single-file skill — push one markdown file as the instruction (no bundle)
python3 create_skill.py --single-file m365-surface-commander/SKILL.md

# Useful flags
python3 create_skill.py --replace      # delete an existing agent of this id first
python3 create_skill.py --share        # set sharingConfig.scope=ALL_USERS after create
```

This mirrors the GE web UI's import flow (create → `files:upload` → get) but uses a plain OAuth
Bearer token (ADC) instead of browser/widget auth. The `agents` resource is undocumented in the
public discovery doc; the authenticated REST endpoints work regardless (verified).

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
