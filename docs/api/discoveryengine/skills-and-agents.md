# Skills & agents on Gemini Enterprise — the live surface vs the published one

This note explains the GE skill layer used by the add-in. The public RPC reference now documents
`StreamAssistRequest.agents_spec`, but the live Gemini Enterprise skill path we have verified is
still the widget-style **`skillsSpec`** field plus a query mention marker. See
`widget-service-skills.md` for the formal widget lifecycle and send-time reference contract.

## The real skill lifecycle

A "skill" is modelled as a GE **agent** with a `skillAgentDefinition`, under the assistant:

```
{assistant} = projects/{num}/locations/{loc}/collections/default_collection
              /engines/{engine}/assistants/default_assistant
```

### Public API mode

Use `discoveryengine.googleapis.com` with OAuth/ADC when the caller has Google IAM permissions on
the assistant resource. This is the preferred API posture for automation, CI, and tenant admin
workflows.

| Operation | Call (authenticated OAuth/ADC bearer) |
|---|---|
| **Create (single-file)** | `POST {assistant}/agents?agentId=<id>` · body `{displayName, description, skillAgentDefinition:{instruction:"<SKILL.md>"}}` |
| **Create (bundle)** | create a shell agent, then `POST /upload/v1alpha/{assistant}/agents/<id>/files:upload?upload_protocol=raw` with `Content-Type: application/zip` — the server unpacks the zip: `SKILL.md` → `instruction`, the rest → `subfiles` |
| **Get / verify** | `GET {assistant}/agents/<id>` → `{name, displayName, state, skillAgentDefinition:{instruction, subfiles[]}}` |
| **Share** | `PATCH {assistant}/agents/<id>?updateMask=sharingConfig` · body `{sharingConfig:{scope:"ALL_USERS"}}` — this is the real "ShareSkill" |
| **Delete** | `DELETE {assistant}/agents/<id>` |

The `agents` CRUD methods are published. The `skillAgentDefinition` payload and the raw
`files:upload` route are the skill-specific layer on top. Note `agents.files.import` (published) is
a *different*, No-Code-only file route; the bundle path used here is the raw
`/upload/v1alpha/.../files:upload`.

**VERIFIED LIVE (saib tenant, WIF token, 2026-07):** the public-API path **works for the whole skill
lifecycle** and needs no widget JWT — `CreateAgent`, `GetAgent`, `UpdateAgent`, `DeleteAgent`, and the
raw `files:upload` bundle route (zip → `SKILL.md` becomes `instruction`, rest become `subfiles`) all
succeed with the signed-in user's federated token. Requirements/caveats from probing:
- `SKILL.md` in an uploaded bundle **must have YAML frontmatter** (`--- name/description ---`) or the
  server returns 400 "Missing YAML frontmatter start delimiter".
- Skill agents need **no `DeployAgent`** (deploy is only for managed/app agents; `:deploy` is 403 here
  and irrelevant to skills). Agents are created `state=PRIVATE` and are still fully usable.
- Only **`ListAgents` (list-all)** and **sharing** (`sharingConfig` PATCH) are denied for this
  principal — minor: we address skills by their known resource id (from `.env`), and PRIVATE is fine.
- The earlier `PERMISSION_DENIED` observation was specifically `discoveryengine.agents.list` (list-all)
  and does **not** apply to create/get/update/delete/upload.

**This makes skill provisioning a boot-time "warm-up" the add-in can do itself** (client-direct, ADR-0001),
not an out-of-band admin/widget step — see `ensureSkillAgent` in `packages/gemini-client/src/agents.ts`,
wired into `composeSession` (`warmUpSkills`). It is idempotent: a cheap `GetAgent` per skill, comparing a
`[rev:<bundle-sha>]` marker in the description, writing only on drift.

### Widget API mode

The Gemini Enterprise web UI uses `content-discoveryengine.googleapis.com` widget calls such as
`widgetListAvailableAgentViews`, `widgetCreateAgent`, `widgetGetAgentView`,
`widgetDeleteAgent`, and a resumable `/upload/v1alpha/.../files:upload` flow. Those calls authorize
against the signed-in user's GE widget permissions (`canRun`, `canView`, `canEdit`, `canDelete`,
etc.) and use a short-lived Vertex AI Search widget JWT.

Use this mode for developer-owned private skills when public IAM is not available. Keep it out of
CI and do not automate browser cookies or XSRF state; copy only the short-lived widget Bearer token
into a local temp file and let the tooling validate issuer/audience/expiry before it sends a
request. The exact local contract is in `widget-service-skills.md`.

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
actually ran — keep using it for provenance.) The query text also carries the GE-style mention
marker, for example `[m365-command-planner](mention://?uri=17573173582293271726)`, because the
widget UI uses those visible mentions to select a skill in natural-language turns.

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
