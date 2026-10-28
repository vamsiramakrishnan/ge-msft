# Surface Commander evaluation (ADR-0008 §10 Phase 4)

The question this evaluation answers: **does the deterministic preflight helper (`surface_cli`)
actually improve execution quality, or just add latency?** It has two halves.

## Offline (deterministic) — `eval_harness.py`

Model-free. Runs `surface_cli` over a **labeled corpus** (`eval-corpus.jsonl`: valid programs +
intentionally-defective ones across every defect category) and measures:

- **recall** — the fraction of seeded defects the helper flags (it should be 100%);
- **false-positive rate** — valid programs it wrongly flags (it should be 0%);
- **per-defect-category** detection (unknown-verb, unknown-capability, capability-violation,
  unbound-var, budget, parse-error);
- **program-length** distribution.

```
python3 eval/eval_harness.py          # report + regression gate
python3 eval/eval_harness.py --json   # machine-readable
```

It exits non-zero if any seeded defect is missed or any valid program is flagged — so the helper's
quality cannot silently regress as the language evolves. This is wired into `test_tooling.py`.

## Online (model-in-the-loop) — needs a live engine

The full comparison the recommendation asks for runs the SAME task set under **three conditions** and
compares them:

1. **base model** — no skill, no helper;
2. **Surface Commander** — the skill instructions only;
3. **Surface Commander + helper** — the skill plus the `surface_cli` preflight loop.

Metrics that require live model runs (not computable offline):

| Metric                                     | Why it needs a model                        |
| ------------------------------------------ | ------------------------------------------- |
| first-pass parse success                   | depends on what the model emits             |
| capability-violation rate                  | model behavior under the injected signature |
| average repair turns                       | counts the model's self-correction loops    |
| unsupported-verb rate                      | model reaching outside the signature        |
| effect-budget rejection rate               | model over-generating effects               |
| approval-preview correction rate           | model vs the previewed plan/DAG             |
| helper latency                             | wall-clock of the preflight call            |
| % helper invocations finding a real defect | precision in the live loop                  |

The harness for these is `test_skill.py` (the live multi-surface runner) driven over the task set
under each condition; this directory's offline harness supplies the deterministic backstop (recall +
false-positive rate) that must hold regardless of the model. Run the online evaluation before claiming
the helper earns its latency — the offline gate only proves it _can_ catch the defects, not that the
model _benefits_ from it in the loop.

## Deterministic matrix — `matrix_harness.py`

The three-condition comparison above, in a form that runs without a model. `eval/matrix/*.jsonl`
holds one row per **task × condition** (`base` · `commander` · `commander+cli`) with the turn
sequence the model produced — recorded from a live run where one exists, otherwise a synthetic but
representative sequence (a base-model turn that uses a verb the language lacks, a commander turn
that is right first time, a `commander+cli` turn whose preflight catches a scope or arity defect
before it reaches the runtime). Every row carries an `expect` block; the harness replays each turn
through `parse_commands` (the grammar, with `def … end` recipes grouped as the runtime groups them)
and `surface_cli.analyze` (scope, type and column laws, recipe arity, budgets) and reports, per
condition:

| metric                        | meaning                                                              |
| ----------------------------- | -------------------------------------------------------------------- |
| first_pass_parse_success      | the FIRST turn is a complete, error-free `cmd` program               |
| capability_violation_rate     | a turn used a verb or `/kind` outside the row's signature            |
| unsupported_verb_rate         | a turn used a verb the language does not have                        |
| effect_budget_rejection_rate  | a turn exceeded the manifest's per-turn budget (`limits`)            |
| avg_repair_turns              | turns after the first until a valid program (capped at the runtime's two; `repair_exhausted_rate` counts rows that never got there) |
| mean_program_length           | non-blank lines of the final program                                 |
| helper_real_defect_share      | `commander+cli` only: rows where the preflight found a defect the grammar alone did not |

The budget numbers come from the generated language manifest (`surface_cli.budget.DEFAULT_LIMITS`
← `m365-cli-1.0.json` `limits`), never a local constant, so the matrix and the runtime gate agree.

```
python3 eval/matrix_harness.py          # table + gate (exit 1 on a failed expectation / ship bar)
python3 eval/matrix_harness.py --json
```

**Gate.** A row whose observed metrics differ from its `expect` block fails the run, and so does
the **ship bar**: `commander+cli` must not have lower first-pass success *and* more repair turns
than `commander`, nor more capability violations left unrepaired. `test_tooling.py` and CI run it
next to `eval_harness.py`. `matrix-report.json` is a generated artifact (git-ignored) rewritten on
every run.

**Adding a task.** Append three rows to `eval/matrix/conditions.jsonl` with the same `task`,
`surface` and `capabilities` (the injected signature) and one row per condition. `turns` is the
list of model replies in order (fenced ```cmd blocks; prose turns count as a failed first pass);
`expect` holds the honest outcome: `first_pass`, `repair_turns`, `unsupported`, `violation`,
`budget`, and for `commander+cli` rows `cli_found_defect`. Record what the model actually did —
the point of the matrix is to show where the skill and the helper change behaviour, not to flatter
them. The corpus covers `each … end` loops (`excel-each-flags`) and recipe definitions and calls
(`word-recipe-reply`, `excel-recipe-total`) as well as single effects, pipelines, analysis
materialization and budget overflow.
