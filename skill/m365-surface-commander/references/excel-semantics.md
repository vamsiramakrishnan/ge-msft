---
title: Excel Semantics
kind: reference
skill: m365-surface-commander
surface: excel
topics: [ranges, tables, formulas, charts, formatting, compute]
load_when: The active surface is Excel or a plan mentions ranges, formulas, tables, charts, or workbook analysis.
---

# Excel semantics

Load this when the active surface is Excel. The cross-surface capability table is in
[capability-map.md](capability-map.md); this is the Excel-specific reading/writing model.

**Selectors** are A1 / `Sheet!A1:G9` / a named range. `read` takes a selector; everything writes by
address. Read the smallest range that answers the request, not the whole sheet. `list table`,
`properties <ref>`, and `open <ref>` expose workbook tables and named ranges as typed refs; resolving
one must use that exact ref, not the current selection.

**Writing is address-anchored** — `set <cell> <value|=formula>` lands an exact cell. Prefer a native
formula (`=SUMIF(...)`) or a value you read over a guessed number. For a multi-cell write, choose the
one bulk shape that matches the source of the data:

- `grid <range> = "a\tb\nc\td"` for a known literal rectangle such as a mock schedule, seed table,
  or CSV-shaped result. This is one reviewable `write-cells` effect.
- `spill <origin> = ($rows)` for a table produced by `read | filter | sort | select | head`; compute
  the resulting range from the origin and table size, then point `table`/`chart`/`cf` at that range.

Do not emit dozens of scalar `set` commands for one logical table.

**Formula safety:** a `=`-prefixed value is evaluated as a formula, but model/host-derived text is
screened — `WEBSERVICE`/`DDE`/external-reference/active-content formulas are refused, not evaluated.
Don't try to route untrusted content through a formula.

**Composition is Excel's strength:** `read … | filter … | sort … | head N` derives a table; `spill`
writes it as a grid at an origin cell; `table`/`chart`/`cf` then consume the **derived range**
(`spill Report!A1` of a 10×2 table → `Report!A1:B11`). The dependent effects depend on the spill —
see [planning-normal-form.md](planning-normal-form.md) and the `top-n-report` pattern.

## Progressive disclosure for advanced Excel

Start cheap: `tables`, `list range`, `properties <ref>`, and bounded `read <range>` before touching
cells, charts, pivots, or worksheets. Escalate to
`context analytical full-scope upload-preferred code-execution-preferred` only when bounded reads
cannot safely answer the task: workbook-scale reconciliation, many sheets, pivot/chart data shaping,
unknown columns, heavy aggregation, or a generated rectangle too large to review inline. The
`context` line is still an OBSERVE request; wait for the host's structured file/code-execution result
and never invent file ids, code outputs, or workbook contents.

Advanced Office.js actions are slash capabilities, not core algebra. Use them only when the current
surface advertises the exact kind or live help exposes it; otherwise fail closed or use an equivalent
core command if it is genuinely equivalent and within scope.

| Capability | Use when | Guardrail |
| ---------- | -------- | --------- |
| `/insert-pivot` | create a native pivot from a confirmed table/range | read/properties the source first; if grouping or measures require workbook-scale analysis, request `context analytical ... code-execution-preferred` before emitting |
| `/sort-range` | reorder an existing worksheet range in place | do not use for derived top-N reports; use pure `sort | head | spill` when the source must remain untouched |
| `/filter-range` | apply a native worksheet filter/view to an existing range | do not substitute for pure `filter` when the user asked for an extracted report |
| `/manage-worksheet` | add, rename, move, hide, protect, or delete sheets | confirm exact sheet identity/name from `list`/`properties`; destructive or ambiguous sheet changes fail closed |
| `/format-chart` | change an existing chart's style, axes, labels, colors, or data binding | resolve the chart object/range first; if no chart exists, use `chart` to create one instead of formatting a guessed target |

Slash commands are effect terminals: they do not pipe, do not produce table values, and count toward
the approval/effect budget. If a slash kind is absent, stale, or unsupported, do not approximate it by
rewriting cells unless the user asked for that explicit fallback and the target range is known.

**Core verbs:** `set`, `grid`, `format`, `comment` (cell-anchored), `reply`, `table`, `chart`, `cf`,
`spill`.
**Specialized (`/`):** capability-gated host-native terminals such as `/insert-pivot`,
`/sort-range`, `/filter-range`, `/manage-worksheet`, `/format-chart`, and `/set-entity-card` (linked
data type — typings-limited; may degrade).

## Chart selection rubric

Always read the live range first and classify the question before choosing `chart`. A "nice chart"
request is not a chart type; infer the visualization only when the data and goal make it obvious.
If two materially different readings are plausible, stop with a clarifying question through the
planner instead of creating a misleading chart.

Use these defaults:

- ranked categories or long labels -> `bar`; sort or top-N first when the source has many categories.
- short category labels by one measure -> `column`.
- dates, times, or ordered periods -> `line`; use `column` only when the periods are few and discrete.
- two numeric variables where relationship/correlation matters -> `scatter`.
- parts of one meaningful total, no negatives, and <=6 categories -> `pie`; otherwise use `bar`.
- cumulative/stacked trend over ordered periods -> `area`, only when the cumulative reading is intended.
- schedules, calendars, Gantt-like grids, or text-heavy matrices -> do **not** chart the raw grid.
  First derive a small summary table such as `Activity | Hours`, `Day | Busy blocks`, or
  `Task | Duration`; write it with `grid`/`spill`, then chart that summary range.
- single numeric column without labels -> read adjacent label/header columns or ask what labels to use.
- selection with mostly blanks/text -> summarize counts or durations first; never chart sparse cells
  just because they are selected.

For schedule visualizations, prefer this shape:

```cmd
read 'Daily schedule'!B1:I54
grid 'Daily schedule'!K6:L18 = "Activity\tHours/Week\nDeep Work\t10\nLunch\t7"
chart bar 'Daily schedule'!K6:L18 title="Weekly Hours by Activity" series=columns
```

The summary table is part of the user-visible workbook state unless the user asked only for analysis.
If the host advertises `/manage-worksheet`, you may create a dedicated summary/chart sheet; otherwise
place the summary in an empty side range after reading/properties confirm that it is safe.

**Gotchas:** comments anchor on a cell you've read; conditional formats apply to a whole range in one
effect (prefer one `cf` over N comments for "highlight"); chart `type` is one of
`column|bar|line|pie|scatter|area`. If the user selects a single numeric column and asks for a chart,
read adjacent label columns before choosing the chart range; do not chart an unlabeled selection
unless the user explicitly asked for that exact range. `chart` inserts an Office-native chart over an
existing verified range; it does not create a worksheet, chart sheet, PNG, matplotlib artifact, or
hosted-code image. If the user asks for a new chart sheet/workspace, use `/manage-worksheet` only
when that slash capability is advertised; otherwise create the chart in-place or stop with `done`.
