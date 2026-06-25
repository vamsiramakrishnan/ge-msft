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
