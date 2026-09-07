# Workflow recipes

The Excel pane presents three guided workflows. Choose one, enter source ranges, preview the result,
then choose a destination and review the exact write. Source chips and column labels come from
captured artifacts. Column numbers in the UI start at 1 within the source range; SDK indices start at 0.

| Recipe ID (version 1) | Parameters | Result |
| --- | --- | --- |
| `reconcile-tables` | Two ranges, key/amount columns, optional currency columns, fixed currency and tolerance | Matched, variance, unpaid, unallocated and invalid groups |
| `duplicate-rows` | Source range, key column, case sensitivity | Nonblank duplicate keys, occurrence count and additional row count |
| `summarize-by-group` | Source range, group/amount columns and currency | Exact totals by group/currency, counts and invalid status |

Column mapping and optional settings expand on demand. Captured headers replace numeric inputs with
named column choices. Explicitly saved settings stay on this device and are validated before reuse.
They never save result data, credentials or approval decisions. Every new preview reads current sources.

An edited form cannot write its old preview. Empty results offer no write; partial results require a
narrower source. Read-only profiles still permit preview when analysis is available. A write is shown
as verified only after verified readback with healthy recovery. Interrupted or unknown outcomes stay
uncertain across preview refresh; inspect Recovery & undo before another overlapping write.

## SDK and CLI

```ts
import {
  compileWorkflowRecipe,
  compileAnalysisProgram,
  inspectAnalysisProgram,
  listWorkflowRecipes,
} from '@ge/runtime';

const catalog = listWorkflowRecipes(); // schemas, UI fields and capability metadata
const program = compileWorkflowRecipe('reconcile-tables', {
  leftRange: 'Invoices!A1:C100',
  rightRange: 'Payments!A1:C100',
  leftKey: 0,
  rightKey: 0,
  leftAmount: 1,
  rightAmount: 1,
  leftCurrency: 2,
  rightCurrency: 2,
  tolerance: '0.01',
});

const dependencies = inspectAnalysisProgram(program);
const cli = compileAnalysisProgram(program);
for await (const event of session.runAnalysisProgram(program)) renderEvent(event);
```

The existing `AssistSession` owns execution; the caller provides `renderEvent`. Omitting `destination`
produces a preview-only program with zero model calls. Adding a destination compiles a guarded
materialization step and requires the existing `approvePlan` callback. The UI deliberately compiles
preview first, then calls `runAnalysis({kind:'materialize', id, destination}, {approvePlan})` for the
reviewed artifact. Capability metadata is descriptive; runtime checks remain authoritative.

`inspectAnalysisProgram()` reports dependencies, layers and independent capture groups. Its
`execution: 'serial'` field is intentional: the current host bridge has individual versioned reads,
not an atomic batch-capture port. Captures and source checks must finish before dependent computation
or approved writes proceed.

```cmd
analyze {"kind":"materialize","id":"$result","destination":"Results!A1","whenNonEmpty":true}
finish when=verified
```

`$result` must already be bound by the program. The conditional skips only a fresh, complete empty
result. A skipped receipt reports zero effects. It does not bypass stale/truncated checks or authorize
a later write. The generated CLI can be checked with the existing Python preflight:

```bash
python3 skill/m365-surface-commander/scripts/surface_cli.py check --surface excel --json < program.cmd
```

## Precision and extension contract

Financial values use `DECIMAL(38,6)`: at most 32 integer and six fractional digits. Store amounts beyond
JavaScript's exact numeric range as decimal text. Malformed, scientific-notation or higher-precision
strings are invalid, not rounded. A group containing invalid amounts has a null total; reconciliation
also has a null variance. Review invalid rows instead of treating them as zero.

To add a recipe, supply a strict versioned input schema, field descriptors, result description and
compiler in `workflow-recipes.ts`. Reuse typed captures, queries and materialization; keep SQL structure
static and admitted references explicit. Query `requiredColumns` provides actionable column checks
before SQL. Add an actual-WASM acceptance case proving output and approval behavior. Behavioral changes
that invalidate saved settings require a new recipe version and an explicit migration policy.

Run `bun run test:command-efficiency` for the local comparison. Run `bun run test:streamassist:modes`
and `bun run test:command-workflows:live` only with the documented tenant configuration. The latter
uses a real provider and compute engine but simulated Office; it cannot certify Office-host latency.
