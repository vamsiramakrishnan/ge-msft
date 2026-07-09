# Invoking agents from the add-in (via StreamAssist)

> **VERIFIED LIVE (saib tenant, 2026-07 — supersedes earlier caveats).** The invocation lever that
> actually routes a skill on the **public** `discoveryengine.googleapis.com:streamAssist` is the
> **`mention://?uri=<agentId>` marker in `query.text`** — confirmed to emit the skill's real output
> (e.g. `invokedSkills=[m365-surface-commander]` + a ` ```cmd ` block). This is exactly what
> `stream-assist.ts` already does. The other selectors do **not** work here: `skillsSpec` is
> accepted-but-ignored, `agentsConfig.agent`/`inlineAgent` are silently ignored (answer comes from
> the default assistant), and `agentsSpec.agentSpecs[].agentId` returns **500**. Those `agentsConfig`/
> `agentsSpec` fields exist only in the GE reference / internal `v1main` serving path, not the public
> endpoint we reach. **Reference a skill by its `mention://?uri=<numericId>` marker; the numeric id
> is the canonical agent resource id (getCard confirms it; displayName is not an id).**

How the client-direct add-in targets a specific Gemini Enterprise agent. For our uploaded GE skills,
mount the skill with `skillsSpec` and prefix the query with the corresponding `mention://` marker.
For ordinary assistant-level routing, point at the **assistant/engine** configured to route to it, or
let the default assistant route.

## What the request can and can't do

- `StreamAssistRequest` fields now include public `agentsSpec`, but this is not the same contract as
  the observed widget private-skill `skillsSpec`.
- Uploaded command-planner/surface-commander skills use `skillsSpec.skills[].name` plus the visible
  `mention://?uri=<agent-id>` marker.
- `actionSpec.actionDisabled` — toggle whether the assistant may take actions this turn.

## Routes to "invoke agent X"

1. **Skill mount for this add-in's planner/executor (recommended for GE skills).** Use
   `skillsSpec.skills[].name` with the full `{assistant}/agents/{agent}` resource and prepend the
   matching mention marker. This is what the web shell tests pin.
2. **Assistant/engine targeting.** Agents are registered on an
   assistant (`engines.assistants.agents.{create,list,get}`). To invoke a specific agent reliably,
   point `cfg.assistant` at the assistant/engine configured for it; an "agent picker" in the UI is
   really an `AssistantPath` selector. Swapping the target is the whole mechanism — no server state.
3. **Default assistant auto-routing.** The default assistant routes by query (and its configured
   agents/tools). Good for the "just ask" path where the user doesn't pick an agent.
4. **A2A direct (special cases only).** You can call an agent's A2A endpoint directly, but that
   **bypasses GE grounding + Model Armor + audit**, so reserve it for non-grounded tools. Prefer
   routing through `streamAssist` so guardrails stay engine-enforced (a hard precondition in
   client-direct — see ADR-0001).

## Registration is an admin task (out of scope for the add-in)

Per Google's docs, a custom agent is: (1) built to the A2A (and optionally A2UI) spec, (2) hosted on
a public endpoint (Cloud Run / Agent Runtime), (3) **registered with Gemini Enterprise by an
administrator** as an A2A agent. The add-in neither registers nor invokes agents out of band — it
selects an assistant target and calls `streamAssist`. See
[Register & manage A2A agents](https://docs.cloud.google.com/gemini/enterprise/docs/register-and-manage-an-a2a-agent)
and [A2UI agents](https://docs.cloud.google.com/gemini/enterprise/docs/a2ui-agents/register-and-manage-an-a2ui-agent).

## Agents that drive UI

A registered agent can answer with **A2UI** (interactive components) instead of plain text; the
add-in renders it and maps its actions to host actuations. See `a2ui.md`. This is how a custom
"Review" or "Redline" agent offers Accept/Reject/Pick-a-rewrite controls that land as tracked changes
in whatever Office app the user is in.
