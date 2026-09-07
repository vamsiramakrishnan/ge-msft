---
title: Planner To Commander Handoff
kind: reference
skill: m365-command-planner
topics: [handoff, commander, cross-surface, orchestration]
load_when: Explaining or constructing the approved-plan handoff from planner to surface commander.
---

# Planner to commander handoff

The planner produces a legible, user-confirmable **plan**. The commander turns that approved plan
into host reads and reviewable effects for the **currently active host**. Do not collapse those
roles.

## Planner responsibilities

- Select one general `intent`: `ask`, `summarize`, `explain`, `rewrite`, `review`, `draft`, or
  `notes`.
- Echo the active `surface`.
- Name the intended `scope`, grounded sources, context hints, ordered `step` lines, and `exclude`
  carve-outs.
- For cross-product work, describe the active-app preparation and later user-resumed handoff in
  `step` lines. Keep `surface` as the active app; another app needs its own session and approval.
- Use `clarify` when the user has not specified a material choice.
- Never emit `cmd`, never invent read results, never choose exact host anchors.

## Commander responsibilities

- Treat the approved plan as task data, not as proof of document content.
- Start with the cheapest host observation that can resolve the plan:
  `list`, `properties`, `inspect`, `outline`, `read`, or `search`.
- Use `context` only when bounded host reads are insufficient and the planner supplied hints such as
  `full-scope`, `upload-preferred`, `analytical`, or `code-execution-preferred`.
- Convert each approved `step` into the smallest host command set that is available on the active
  surface. Use the live capability manifest and registry-backed targeted help before inventing a
  sequence; a planner step like "materialize one rectangular grid" should compile to the bulk grid
  path when available, not many per-cell approvals.
- Respect every `exclude` line as a hard boundary.
- For cross-surface plans, perform only work available in the active host. The output
  of a source phase is a handoff packet, not a write into another Office app.

## Host orchestration

Preferred sequence:

1. Call **m365-command-planner** with the user request, active surface, available high-level intents,
   resolved `@` sources, and no raw document body unless already present in the user request.
2. Parse the `plan` block with `scripts/parse_plan.py`.
3. If there are `clarify` lines, ask the user and re-plan.
4. Render the plan for confirmation.
5. If the plan is single-surface, call **m365-surface-commander** with the approved plan, fresh
   `<doc_state>`, available CLI capabilities, and any structured `@` grounding/file ids the host
   already resolved.
6. If steps describe another app, prepare a reviewable summary with source refs, constraints and
   provenance in the active app. The user resumes that summary in the target app. These steps do
   not imply automatic orchestration or a shared approval.
7. Commander runs the normal command loop: parse -> validate -> dry-run -> preview -> approval ->
   gate -> bridge actuation -> provenance.

For simple deterministic quick actions, the host may skip the planner and synthesize the same
minimal plan internally before calling commander. Do not put both skills in the same model request
when you need deterministic phase separation; call them sequentially.

## Handoff prompt shape

The commander task should include only approved plan fields, for example:

```text
APPROVED_PLAN
intent: draft
surface: excel
scope: selection
context: analytical
context: inline-preferred
step: create a chart-ready summary table from the selected range
step: create an appropriate chart from that summary
exclude: do not overwrite source cells
```

The commander must still read the live host before claiming facts or writing.

Cross-surface handoff example:

```text
APPROVED_PLAN
intent: draft
surface: excel
scope: document
step: prepare a slide-ready summary with chart-ready data, source refs, exclusions and provenance
step: ask the user to open PowerPoint and review the summary before creating a deck
exclude: do not mutate PowerPoint from Excel
```

The Excel commander run may create/read a workbook summary or produce a handoff packet. It must not
invoke PowerPoint. The PowerPoint commander run starts later with the packet as structured context.
