# ADR-0014: Workflow recipes and deterministic execution state

Date: 2026-09-07
Status: Implemented; live tenant and Office acceptance remain release gates.

Known workflows should not need an inference turn to reconstruct a program that the runtime already
understands. Longer independent command requests should carry current execution state and retrieve
prior evidence when needed. The pane should expose these capabilities through meaningful actions,
with the same approval and completion rules as the CLI.

## Decisions

Versioned recipe definitions own parameter schemas, field metadata, result descriptions and compilation.
The initial recipes reconcile two tables, identify duplicate keys, and summarize amounts by group.
UI forms and SDK callers use the same definitions. Recipes compile to the existing `AnalysisProgram`
and `cmd` grammar; there is no second execution engine or delimiter. Dependency inspection exposes
independent captures, while actual bridge execution remains serial.

The pane separates preview from materialization. Preview strips any write destination, runs without
model inference or an approver, and associates the result with the emitted `$result` binding and
normalized recipe inputs. Editing parameters invalidates the displayed result. Writes still require
an explicit destination, fresh sources, exact plan approval, a recovery checkpoint and readback.
Saved settings contain versioned parameters only; they carry neither artifacts nor approval authority.
Effective capabilities distinguish preview availability from write availability.

`whenNonEmpty` is a typed optional materialization condition. Both direct actions and compiled programs
validate source freshness and completeness first. A fresh empty result produces an explicit skipped
receipt with zero effects. Stale or truncated data cannot claim a safe skip. A skip needs no write
capability because no write is prepared.

Sessionless command context defaults to deterministic projection. It pins the original task, current
bindings, artifact schemas, macro references, actual effect outcomes, historical failures and latest
results. Complete observations remain in a bounded task-local journal retrievable with `inspect
state:…`. This is structural projection, not a generated summary. Serialization rejects accessors,
custom JSON conversion, cycles and oversized data before disclosure. Budget overflow stops; it cannot
silently discard constraints or uncertain effects. Explicit transcript compatibility remains available.

Execution state is untrusted data, never a plan, approval or proof of success. Journal references expire
with the task. A previous program does not authorize replay. Recovery independently rejects overlapping
unresolved writes using captured document, worksheet and cell bounds, including new change IDs and
address aliases. A previously verified write followed by a user edit remains historical conflict data;
it cannot lock the range permanently or acquire undo authority over the user's newer changes.

## Consequences and evidence

The familiar CLI stays stable while recipe definitions provide a small extension point for new
workflows. Financial recipes use shared exact-decimal validation: unsafe native numbers fail with
decimal-text guidance, and invalid precision or malformed strings produce invalid groups with null
totals. They cannot silently round or report partial totals as complete.

Actual DuckDB fixtures verify all three recipes with zero model calls and one approval for nonempty
writeback. Projection reduces submitted query bytes by 28.6% in the nine-turn evidence fixture and
increases them by 7.9% in the five-turn short fixture. It is not universally smaller. These are local
encoding measurements with simulated provider/Office adapters, not live latency or billing results.

Rendered React tests are now collected by the main test configuration, including stale parameters,
cancellation, lost artifacts, storage failure, read-only profiles and uncertain write outcomes.
The cloud browser could not reach the local preview, so no pixel-level browser validation is claimed.
The live harness separately reports request mode, private skill routing, model/compute/approval/host
timings and p50/p95. Real tenant execution and Office readback remain explicit release gates.

See [workflow usage](WORKFLOW-RECIPES.md) and [performance measurements](COMMAND-PERFORMANCE.md).
