# Invoking agents from the add-in (via StreamAssist)

How the client-direct add-in targets a specific Gemini Enterprise agent. The short version:
**`v1alpha` `streamAssist` has no `agentsSpec`**, so you don't name an agent in the request body —
you point at the **assistant/engine** configured to route to it, or let the default assistant route.

## What the request can and can't do

- `StreamAssistRequest` fields: `query`, `session`, `userMetadata`, `toolsSpec`, `generationSpec`,
  `actionSpec`. **No `agentsSpec`** — the early-2026 agent-id bug is avoided structurally because the
  field doesn't exist here (it lived in other versions/paths). Routing is the assistant's job.
- `actionSpec.actionDisabled` — toggle whether the assistant may take actions this turn.

## Routes to "invoke agent X"

1. **Assistant/engine targeting (recommended, deterministic).** Agents are registered on an
   assistant (`engines.assistants.agents.{create,list,get}`). To invoke a specific agent reliably,
   point `cfg.assistant` at the assistant/engine configured for it; an "agent picker" in the UI is
   really an `AssistantPath` selector. Swapping the target is the whole mechanism — no server state.
2. **Default assistant auto-routing.** The default assistant routes by query (and its configured
   agents/tools). Good for the "just ask" path where the user doesn't pick an agent.
3. **A2A direct (special cases only).** You can call an agent's A2A endpoint directly, but that
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
