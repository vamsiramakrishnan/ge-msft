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
formula (`=SUMIF(...)`) or a value you read over a guessed number.

**Bulk generated rectangles:** use `grid <range> = "a\tb\nc\td"` for a known rectangular payload
such as a mock schedule, seed table, or CSV-shaped result. This is one reviewable `write-cells`
effect. Do not emit dozens of scalar `set` commands for one logical table.

**Formula safety:** a `=`-prefixed value is evaluated as a formula, but model/host-derived text is
screened — `WEBSERVICE`/`DDE`/external-reference/active-content formulas are refused, not evaluated.
Don't try to route untrusted content through a formula.

**Composition is Excel's strength:** `read … | filter … | sort … | head N` derives a table; `spill`
writes it as a grid at an origin cell; `table`/`chart`/`cf` then consume the **derived range**
(`spill Report!A1` of a 10×2 table → `Report!A1:B11`). The dependent effects depend on the spill —
see [planning-normal-form.md](planning-normal-form.md) and the `top-n-report` pattern.

**Core verbs:** `set`, `grid`, `format`, `comment` (cell-anchored), `reply`, `table`, `chart`, `cf`,
`spill`.
**Specialized (`/`):** `/set-entity-card` (linked data type — typings-limited; may degrade).

**Gotchas:** comments anchor on a cell you've read; conditional formats apply to a whole range in one
effect (prefer one `cf` over N comments for "highlight"); chart `type` is one of
`column|bar|line|pie|scatter|area`. If the user selects a single numeric column and asks for a chart,
read adjacent label columns before choosing the chart range; do not chart an unlabeled selection
unless the user explicitly asked for that exact range.
