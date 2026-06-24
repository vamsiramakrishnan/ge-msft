# Skills & agents on Gemini Enterprise — the live surface vs the published one

This note **corrects and extends** `agent-invocation.md`. That doc concluded "there is no
`agentsSpec`, so you can't name an agent/skill in the `streamAssist` request — routing is the
assistant's job." That is true of the **published** `discoveryengine.v1alpha` discovery doc and
proto (verified: `StreamAssistRequest` = `session, actionSpec, query, userMetadata, toolsSpec,
generationSpec`; no `skills_spec`, no `agentsSpec`; the only skill field is `invokedSkills[]` on
the **response**).

It is **not** the whole live surface. The `engines.assistants.agents.*` CRUD methods **are**
published (they appear in `methods-index.md`: `agents.create/get/list/patch/delete` and
`agents.files.import`). What the published schema **omits** is the skill-specific shape on top of
them — the `skillAgentDefinition` payload on an Agent, the raw `files:upload` endpoint, and the
**`skillsSpec`** field on `streamAssist`. Those three were verified end-to-end against a live engine
(see `skill/` tooling: `create_skill.py`, `test_skill.py`, README "What we learned"). Treat the
published doc as a **subset** of what the endpoint accepts.

## The real skill lifecycle (verified)

A "skill" is modelled as a GE **agent** with a `skillAgentDefinition`, under the assistant:

```
{assistant} = projects/{num}/locations/{loc}/collections/default_collection
              /engines/{engine}/assistants/default_assistant
```

| Operation | Call (authenticated, ADC bearer) |
|---|---|
| **Create (single-file)** | `POST {assistant}/agents?agentId=<id>` · body `{displayName, description, skillAgentDefinition:{instruction:"<SKILL.md>"}}` |
| **Create (bundle)** | create a shell agent, then `POST /upload/v1alpha/{assistant}/agents/<id>/files:upload?upload_protocol=raw` with `Content-Type: application/zip` — the server unpacks the zip: `SKILL.md` → `instruction`, the rest → `subfiles` |
| **Get / verify** | `GET {assistant}/agents/<id>` → `{name, displayName, state, skillAgentDefinition:{instruction, subfiles[]}}` |
| **Share** | `PATCH {assistant}/agents/<id>?updateMask=sharingConfig` · body `{sharingConfig:{scope:"ALL_USERS"}}` — this is the real "ShareSkill" |
| **Delete** | `DELETE {assistant}/agents/<id>` |

This mirrors the GE web UI import flow (create → `files:upload` → get) but with a plain OAuth
bearer token instead of browser/widget auth. The `agents` CRUD is published; the
`skillAgentDefinition` payload and the raw `files:upload` route are the undocumented part — they
work regardless (verified). Note `agents.files.import` (published) is a *different*, No-Code-only
file route; the bundle path used here is the raw `/upload/v1alpha/.../files:upload`.

## Mounting a skill per turn — `skillsSpec` (the load-bearing correction)

To run a specific skill for a turn, reference its agent **resource name** in `streamAssist`:

```jsonc
{
  "query":      { "text": "<the task or the confirmed plan>" },
  "session":    "<engine session>",
  "skillsSpec": { "skills": [ { "name": "projects/…/agents/m365-surface-commander" } ] }
}
```

This is **deterministic** — unlike auto-discovery, the client chooses the skill. So our earlier
"you can only point at an assistant and hope it routes" was too pessimistic: with `skillsSpec`
we mount exactly the skill we want. (`invokedSkills[]` in the response still reports what
actually ran — keep using it for provenance.)

### Reliability notes (from live refinement)

- **Isolate with an empty `toolsSpec`** — no web grounding / data stores leak into a skill test.
- **Inject each available verb's *exact usage* into a per-turn `<capabilities>` block** — the
  single highest-leverage lever against verb/syntax drift. This is the host's job and maps
  directly onto our `CapabilityManifest` (ADR-0002) — the manifest must render *usage*, not just
  verb names.
- A `skillsSpec`-layered skill **tends to answer turn 1 in prose**, then emit correct output
  after a re-prompt. The host must: re-prompt on a no-command/no-plan turn, and **not** honour
  `done` if that same block had parse errors.

## How this lands in the `/` + `@` design

The corrected mechanism makes the command pane simpler and deterministic:

- **`/` verb → mount the right skill via `skillsSpec`.** Two skills, mounted by the host:
  - **`m365-command-planner`** (front door): for complex free-text requests, emits a structured
    ` ```plan ` block → host renders it for one-tap confirm (the legibility gate) → dispatch.
  - **`m365-surface-commander`** (executor): takes the confirmed plan + live document snapshot
    and emits ` ```cmd ` lines → runtime parses → `gate()` → tracked change / cell / staged draft.
  - Route by complexity: verb + mentions with little free text → skip the planner, go straight to
    the executor; free text with constraints/exclusions → planner first.
- **`@` mention → grounding** still maps to the real published fields (`query.parts`,
  `toolsSpec.vertexAiSearchSpec.dataStoreSpecs`, `fileIds`) — unchanged and verified.
- Both skills can be mounted in one session (`skillsSpec.skills:[planner, executor]`); the
  observed prose-first turn naturally becomes the plan, which we confirm, then the session
  continues into the `cmd` loop.

## Parity tasks (skill ⇄ workspace) — keep these in lockstep

The skill bundles in `skill/` and our TypeScript runtime are two implementations of the **same**
grammar and wire reader. The TS side is authoritative (it applies changes); the Python is the
skill's self-check. They must agree:

| Skill artifact (`skill/…`) | Workspace counterpart | Invariant |
|---|---|---|
| `m365-surface-commander/scripts/parse_commands.py` | `runtime` command parser + `contracts` CommandGrammar (ADR-0004) | same verbs, selectors, corrective-error wording |
| `m365-surface-commander/references/capability-map.md` | per-bridge `CapabilityManifest` (ADR-0002) | the manifest renders each verb's **exact usage** into `<capabilities>` |
| `m365-command-planner/scripts/parse_plan.py` | `contracts` `CommandPlan` schema (new) | same keywords; `intent ∈ IntentSchema`; `needs_clarification` blocks dispatch |
| `de_stub.py::read_response` | `gemini-client` `streamAssist` reader | reassemble streamed tokens; split text vs thoughts vs `textGroundingMetadata` citations vs `inlineData` suggestions; survive split fences |

## Status of the original third-party "Skill Service" claim

A separate **Vertex/ADK Skill Registry** also exists on `aiplatform.googleapis.com/v1beta1`
(`CreateSkill`/`GetSkill`/`RetrieveSkills`) — that is a *different* system from the GE
`{assistant}/agents` path we use here. Our add-in uses the **GE agents + `skillsSpec`** path
(above), which is the one verified to drive `streamAssist`. The earlier user-supplied method list
(`UploadSkill`/`SearchSkill`/`DownloadSkill`/`skills_spec = 28`) was directionally describing this
real-but-undocumented capability; the exact names/field number differ from both live systems.
