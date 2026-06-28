---
title: Plan Grammar
kind: reference
skill: m365-command-planner
topics: [plan-block, keywords, cross-surface, clarification]
load_when: Exact plan keyword rules, cross-surface fields, or parser-compatible examples are needed.
---

# Plan grammar (full reference)

A plan is one fenced ` ```plan ` block of flat keyword lines. It is a **structured
intention**, not commands — the executor (`m365-surface-commander`) turns each step into the
actual command after reading the live document. Keep the plan small and legible: it is shown
to the user for a one-tap confirm before anything runs.

## Keywords

| Keyword      | Repeatable | Required | Meaning                                                                                                                                                                          |
| ------------ | :--------: | :------: | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `intent`     |     no     |   yes    | The general verb, one of `ask`, `summarize`, `explain`, `rewrite`, `review`, `draft`, `notes`. Use `ask` (custom prompt) if none fits.                                           |
| `surface`    |     no     |   yes    | The active app where the workflow starts: `word`, `excel`, `powerpoint`, `onenote`, `outlook`, `teams`.                                                                         |
| `scope`      |     no     |    no    | Where the first phase applies — one of `selection`, `document`, `range`, `section`, `comment`, `this-item`. A plain ref may follow.                                             |
| `ground`     |    yes     |    no    | A pinned `@`source this plan relies on, by verbatim title. Must correspond to a source the host supplied.                                                                        |
| `context`    |    yes     |    no    | Context-construction hint: `incremental`, `inline-preferred`, `reference-preferred`, `upload-preferred`, `code-execution-preferred`, `analytical`, or `full-scope`.              |
| `workflow`   |     no     |    no    | `single-surface` or `cross-surface`. Omit or use `single-surface` for normal one-host work.                                                                                      |
| `source`     |    yes     | cross-surface | Source app/scope, starting with a surface, e.g. `excel document`.                                                                                                           |
| `target`     |    yes     | cross-surface | Target app/scope, starting with a surface, e.g. `powerpoint deck`.                                                                                                         |
| `phase`      |    yes     | cross-surface | A per-host phase, starting with a surface, e.g. `excel prepare the handoff packet`.                                                                                        |
| `handoff`    |    yes     | cross-surface | The user-visible artifact passed between phases: summary data, refs, draft text, constraints, provenance, next action.                                                     |
| `step`       |    yes     | yes (≥1) | One intention, in order. Executor-shaped but natural language. One reviewable change per step.                                                                                   |
| `exclude`    |    yes     |    no    | An explicit carve-out — something to leave unchanged.                                                                                                                            |
| `clarify`    |    yes     |    no    | A question to ask the user before executing. Any `clarify` line blocks dispatch until resolved.                                                                                  |
| `confidence` |     no     |    no    | `high` / `medium` / `low` — your read of how well-specified the request is.                                                                                                      |

`plan` and `end` lines are optional brackets; the fence itself delimits the block. Lines
starting with `#` are comments. Unknown keywords are reported back as a corrective error.

## The verb is general; scope and ground are orthogonal

The seven verbs are **general capabilities**, identical on every surface — the verb says
WHAT, `scope` says WHERE, `ground` says what it is grounded on. Do not smuggle a surface or a
task into the verb (no `regen-clause`, `draft-slides`, `synthesize`, `meeting-notes`,
`resolve-comment` — those are scopes/closures of the seven):

| Verb        | Means                                                    | Route         |
| ----------- | -------------------------------------------------------- | ------------- |
| `ask`       | a custom free-text prompt / grounded chat over the scope | chat (read)   |
| `summarize` | condense the scope                                       | chat (read)   |
| `explain`   | clarify the scope in plain language                      | chat (read)   |
| `rewrite`   | apply any instruction to the scope → a reversible edit   | write (gated) |
| `review`    | whole-scope pass → N findings → N gated annotations      | annotation    |
| `draft`     | generate new material (slides, page, reply, column)      | write (gated) |
| `notes`     | transcript → live notes + action items (Teams)           | annotation    |

`rewrite` is the load-bearing generalization: "tighten", "make formal", "rewrite to match
the policy" are all `rewrite` + a free-text `step`, compiling to whatever reversible write
the scope×surface affords (a Word tracked change, an Excel cell, a slide-body replace).
`rewrite scope comment <id>` is how a comment thread is actioned (the old `resolve-comment`).

## How `step` maps per surface

A step is phrased so the executor can realize it with one command on that surface:

| Surface    | A step becomes…                              | Executor verb(s)                    |
| ---------- | -------------------------------------------- | ----------------------------------- |
| Word       | a tracked change or a comment                | `suggest`, `comment`, `reply`       |
| Excel      | a cell value/formula, a format, or a comment | `set`, `format`, `comment`, `reply` |
| PowerPoint | an inserted slide                            | `slide`                             |
| OneNote    | an appended page                             | `page`                              |
| Outlook    | a staged reply or a new draft                | `mail`, `compose`                   |
| Teams      | a staged channel post                        | `post`                              |

Write the intention, not the command: _"rewrite the SLA figure to 99.9% as a tracked change"_
→ the executor reads the clause, finds the exact text, and emits
`suggest "…99.5%…" => "…99.9%…"`. Keep one change per `step` so each maps to one previewed,
approved, recorded edit.

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
- an exclusion conflicts with a step.

A plan carrying any `clarify` line is surfaced to the user as a question; the host re-plans
with the answer rather than dispatching a guess to the executor.

## Cross-surface workflow grammar

Use `workflow cross-surface` when the request explicitly spans apps, for example Excel workbook →
PowerPoint deck, Outlook thread/attachment → Excel analysis, Teams transcript → Word notes, or Word
report → PowerPoint summary.

Cross-surface does **not** mean a single host may mutate another host. It means the host creates a
typed handoff packet after the active-surface phase, then the user resumes in the target app.

Required shape:

```
workflow cross-surface
source   <surface> <scope>
target   <surface> <scope>
phase    <surface> <intention for that host>
phase    <surface> <intention for the next host>
handoff  <artifact contents and constraints>
```

The `surface` scalar remains the active starting app. A `phase` line must start with a surface. A
`handoff` should name enough to be auditably replayed: source refs, generated outline/table/draft,
constraints, exclusions, provenance, and next action.

Bad:

```
step Excel writes the PowerPoint deck directly
```

Good:

```
phase excel create a slide-ready handoff packet from the workbook analysis
phase powerpoint create an executive deck from the approved handoff packet
handoff slide outline, chart-ready table, source refs, exclusions, provenance
```

## The fenced block

Emit exactly one fenced ` ```plan ` block per turn. Only its contents are parsed; any text
outside it is ignored. A turn with no `plan` block is treated as a prompt to try again, not an
error.
