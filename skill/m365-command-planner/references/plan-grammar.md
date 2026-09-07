---
title: Plan Grammar
kind: reference
skill: m365-command-planner
topics: [plan-block, keywords, cross-surface, clarification]
load_when: Exact supported plan keywords, cross-surface scope boundaries, or parser-compatible examples are needed.
---

# Plan grammar (full reference)

A plan is one fenced ` ```plan ` block of flat keyword lines. It is a **structured
intention**, not commands — the executor (`m365-surface-commander`) turns each step into the
actual command after reading the live document. Keep the plan small and legible: it is shown
to the user for a one-tap confirm before anything runs.

## Keywords

| Keyword      | Repeatable |   Required    | Meaning                                                                                                                                                             |
| ------------ | :--------: | :-----------: | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `intent`     |     no     |      yes      | The general verb, one of `ask`, `summarize`, `explain`, `rewrite`, `review`, `visualize`, `draft`, `notes`. Use `ask` (custom prompt) if none fits.                 |
| `surface`    |     no     |      yes      | The active app where the workflow starts: `word`, `excel`, `powerpoint`, `onenote`, `outlook`, `teams`.                                                             |
| `scope`      |     no     |      no       | Where the first phase applies — one of `selection`, `document`, `range`, `section`, `comment`, `this-item`. A plain ref may follow.                                 |
| `ground`     |    yes     |      no       | A pinned `@`source this plan relies on, by verbatim title. Must correspond to a source the host supplied.                                                           |
| `context`    |    yes     |      no       | Context-construction hint: `incremental`, `inline-preferred`, `reference-preferred`, `upload-preferred`, `code-execution-preferred`, `analytical`, or `full-scope`. |
| `step`       |    yes     |   yes (≥1)    | One intention, in order. Executor-shaped but natural language. One reviewable change per step.                                                                      |
| `exclude`    |    yes     |      no       | An explicit carve-out — something to leave unchanged.                                                                                                               |
| `clarify`    |    yes     |      no       | A question to ask the user before executing. Any `clarify` line blocks dispatch until resolved.                                                                     |
| `confidence` |     no     |      no       | `high` / `medium` / `low` — your read of how well-specified the request is.                                                                                         |

`plan` and `end` lines are optional brackets; the fence itself delimits the block. Lines
starting with `#` are comments. Unknown keywords are reported back as a corrective error.

## The verb is general; scope and ground are orthogonal

The general verbs are **surface-agnostic capabilities** — the verb says
WHAT, `scope` says WHERE, `ground` says what it is grounded on. Do not smuggle a surface or a
task into the verb (no `regen-clause`, `draft-slides`, `synthesize`, `meeting-notes`,
`resolve-comment` — those are scopes/closures of these verbs):

| Verb        | Means                                                    | Route         |
| ----------- | -------------------------------------------------------- | ------------- |
| `ask`       | a custom free-text prompt / grounded chat over the scope | chat (read)   |
| `summarize` | condense the scope                                       | chat (read)   |
| `explain`   | clarify the scope in plain language                      | chat (read)   |
| `rewrite`   | apply any instruction to the scope → a reversible edit   | write (gated) |
| `review`    | whole-scope pass → N findings → N gated annotations      | annotation    |
| `visualize` | create a chart, visual summary, or chart-ready table     | write (gated) |
| `draft`     | generate new material (slides, page, reply, column)      | write (gated) |
| `notes`     | transcript → live notes + action items (Teams)           | annotation    |

`rewrite` is the load-bearing generalization: "tighten", "make formal", "rewrite to match
the policy" are all `rewrite` + a free-text `step`, compiling to whatever reversible write
the scope×surface affords (a Word tracked change, an Excel cell, a slide-body replace).
`rewrite scope comment <id>` is how a comment thread is actioned (the old `resolve-comment`).

## How `step` maps per surface

A step is phrased so the executor can realize it with a bounded capability on that surface. It is
not a CLI command, but it should be close enough to the capability vocabulary that the commander can
choose the right command after reading live host content.

| Surface    | Step should name…                                                                                                                                       |
| ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Word       | tracked rewrite; anchored comment/reply; style application; inserted table/hyperlink/content control; bounded find/replace.                             |
| Excel      | one rectangular grid/table materialization; formulas; range format; comment/reply; native table; chart/pivot summary; sort/filter; worksheet operation. |
| PowerPoint | slide creation; selected shape/text update; image/table/chart-ready insertion; slide layout; shape/text formatting; deck handoff import.                |
| OneNote    | page append; page title; outline/rich text block; note tag; explicit section creation.                                                                  |
| Outlook    | staged reply or new draft; body/subject/recipients; attachment; categories; calendar draft. Never automatic send.                                       |
| Teams      | staged post/card; transcript notes/actions; thread/channel handoff; estate/Graph write only when explicitly requested and gated.                        |

Write the intention, not the command: _"rewrite the SLA figure to 99.9% as a tracked change"_
→ the executor reads the clause, finds the exact text, and emits
`suggest "…99.5%…" => "…99.9%…"`. Keep one change per `step` so each maps to one previewed,
approved, recorded edit.

Bulk data entry is one reviewable change when it is one coherent table/grid. For example, a request
to fill a blank weekly schedule should plan:

```
step materialize the requested weekly schedule as one rectangular grid in the existing table
```

It should not plan dozens of per-cell steps. The commander decides whether to use grid, spill, table,
or smaller writes based on the live capability manifest and host limits.

For `visualize`, plan the question the chart should answer, not just "make a chart." If the source is
a schedule/calendar/text matrix, the first step should be to derive a chart-ready summary table
(for example, hours by activity, hours by day, or task duration), then create a native chart from that
summary. If the user's intended metric is unclear, emit `clarify` with compact options instead of
guessing a chart type.

## Scope, grounding, exclusions

- **`scope`** narrows where the executor reads/acts. Leave it off to mean "the whole open item".
- **`ground`** names the sources the plan depends on; the host has already pinned them as
  `@`-mentions and mapped them to real `streamAssist` fields (`query.parts`,
  `dataStoreSpecs`). Only echo the ones you actually use.
- **`context`** names how much material the runtime should consider attaching. It is a hint, not an
  action: the host still decides whether the source is inlined, referenced, or uploaded as a session
  context file for hosted code execution.
- **`exclude`** is a hard carve-out the executor must respect — list anything the free text
  said to leave alone.

## Context strategy

Use these hints sparingly and compositionally:

- `incremental` — read live host slices lazily through the executor (`outline`, `read`, `search`).
- `inline-preferred` — selected item/range/thread is small enough to inline as context.
- `reference-preferred` — pinned/indexed sources should be referenced rather than copied.
- `upload-preferred` — whole file/deck/thread/workbook is likely needed and too large to inline.
- `code-execution-preferred` — hosted Python would help compute/pivot/chart/validate/reconcile.
- `analytical` — the task is data-analysis heavy.
- `full-scope` — the whole open artifact is needed, not just the current selection.

Examples: Excel workbook anomaly detection usually emits `analytical`, `full-scope`,
`upload-preferred`, and `code-execution-preferred`; Word selected-clause rewrite usually emits
`inline-preferred` or `incremental`; Teams full-transcript action extraction emits `full-scope` and
possibly `analytical`.

## When to `clarify`

Emit `clarify` (and keep the plan minimal) when a material choice can't be inferred:

- the request names a standard/policy ambiguously (which control? which version?),
- the scope is unclear (which section/range?),
- the action could mean two different changes,
- the user asks for a chart/visualization but the data supports multiple valid summaries
  (for example, hours by activity vs hours by day vs meeting/focus split),
- an exclusion conflicts with a step.

A plan carrying any `clarify` line is surfaced to the user as a question; the host re-plans
with the answer rather than dispatching a guess to the executor.

## Cross-surface intentions

The runtime parses plans for the active `surface`. It has no cross-surface keyword extension.
Express the current work and review boundary using supported steps:

```text
surface excel
step Prepare a slide-ready summary table and outline from the workbook, with source refs and exclusions.
step Ask the user to open PowerPoint and review the handoff before creating slides there.
exclude Do not mutate PowerPoint from Excel.
```

Use `clarify` if the destination or material scope is unknown. A later app requires its own active
session and review; describing that intention does not start another runtime or carry approval.

## The fenced block

Emit exactly one fenced ` ```plan ` block per turn. Only its contents are parsed; any text
outside it is ignored. A turn with no `plan` block is treated as a prompt to try again, not an
error.
