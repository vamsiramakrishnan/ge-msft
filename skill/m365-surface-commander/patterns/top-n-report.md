---
title: Pattern - Top N Report
kind: pattern
skill: m365-surface-commander
surface: excel
topics: [top-n, table, chart, spill]
load_when: A task asks for ranked rows, top/bottom items, summary table, or chart from spreadsheet data.
---

# Pattern: top-N report

A reasoning template, **not** a command. Read it for the shape, then write the program in the
turn's actual algebra — never treat the pattern name as a verb.

**Intent:** "show the top N <thing> by <measure>, as a table and a chart."

**Preconditions**

- the source supports `read`;
- the target supports `spill`;
- optionally `table` / `chart` are in this turn's signature.
- if the workbook is too large or the measure/grouping is ambiguous, request
  `context analytical full-scope upload-preferred code-execution-preferred` before deriving rows.

**Pure core (OBSERVE → DERIVE)**

```
read <source>
  → filter   (optional — scope to the subset asked for)
  → sort      <measure> desc
  → head      N
  → select    <output columns>
```

Bind it: `let $top = read … | filter … | sort <measure> desc | head N | select <cols>`.

**Effect core (EFFECT)**

```
spill  <origin>            ($top)        # materialize the derived table as a grid
table  <derived-range>     [headers]      # the range = origin × size of $top (computable)
chart  <type> <derived-range> title="…"  # over the SAME derived range
```

`table` and `chart` consume the **derived range** (computed from the spill origin and `$top`'s size),
so both depend on the `spill`. The dependency is inferred — never write it by hand.

**Worked example**

```cmd
let $top = read Sales!A1:D5000 | filter Quarter=Q3 | sort Revenue desc | head 10 | select Region,Revenue
spill Report!A1 = ($top)
table Report!A1:B11 headers
chart column Report!A1:B11 title="Top regions by revenue"
done
```

**Failure rule:** `table`/`chart` depend on `spill`. If the `spill` is rejected at approval, the
dependent effects are skipped — they have nothing to bind to.

**Advanced Excel variants:** `/sort-range` and `/filter-range` mutate the source range; they are not
the right tool for a top-N report unless the user explicitly asked to sort/filter the worksheet in
place and those slash capabilities are advertised this turn. `/insert-pivot` may be the right native
artifact for a pivot-table request, but only after the source table/range and measures are confirmed;
otherwise keep the pure `read | filter | sort | head | spill` flow.

**Anti-patterns**

- charting the raw source (`Sales!…`) instead of the derived range — you'd chart 5000 rows, not 10;
- `head` before `sort` — the wrong 10 rows;
- using hidden cells to stage `$top` instead of `spill`ing it where it belongs.
