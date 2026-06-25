#!/usr/bin/env python3
"""
eval_harness.py — the DETERMINISTIC half of the ADR-0008 §10 Phase-4 evaluation.

It measures the one question that decides whether the preflight helper earns its latency: **does
`surface_cli` catch the defects it should (recall), without crying wolf on valid programs
(false-positive rate)?** It runs the helper over a labeled corpus (valid + intentionally-defective
programs) and reports precision/recall + a per-defect-category breakdown + program-length stats.

This is the offline, model-free slice. The full three-condition comparison (base model · Surface
Commander · Surface Commander + helper) needs LIVE model runs — see eval/README.md for that protocol
and the online metrics (repair turns, approval-preview corrections) it adds. This harness does NOT
call a model or any host; it only exercises the deterministic checker.

Usage:
  python3 eval/eval_harness.py            # report + regression gate (exit 1 if a seeded defect slips)
  python3 eval/eval_harness.py --json
"""

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "m365-surface-commander" / "scripts"))
import surface_cli  # noqa: E402

CORPUS = HERE / "eval-corpus.jsonl"


def _flagged(case) -> bool:
    """The helper's verdict: does it flag this program (an error or a budget overrun)?"""
    result = surface_cli.analyze(case["program"], set(case.get("capabilities") or []) or None)
    limits = {
        "max_effects": case.get("max_effects", 8),
        "max_reads": case.get("max_reads", 8),
        "max_cells": case.get("max_cells", 10000),
    }
    return bool(result["errors"]) or surface_cli._budget_exceeded(result, limits)


def evaluate():
    cases = [json.loads(l) for l in CORPUS.read_text(encoding="utf-8").splitlines() if l.strip()]
    defective = [c for c in cases if c["label"] == "defective"]
    valid = [c for c in cases if c["label"] == "valid"]

    caught = [c for c in defective if _flagged(c)]
    false_pos = [c for c in valid if _flagged(c)]

    by_defect = {}
    for c in defective:
        d = c.get("defect", "unspecified")
        by_defect.setdefault(d, [0, 0])
        by_defect[d][1] += 1
        if _flagged(c):
            by_defect[d][0] += 1

    lengths = [len([ln for ln in c["program"].splitlines() if ln.strip()]) for c in cases]
    return {
        "total": len(cases),
        "defective": len(defective),
        "valid": len(valid),
        "recall": len(caught) / len(defective) if defective else 1.0,
        "false_positive_rate": len(false_pos) / len(valid) if valid else 0.0,
        "false_positives": [c["name"] for c in false_pos],
        "missed": [c["name"] for c in defective if not _flagged(c)],
        "by_defect": {d: f"{hit}/{tot}" for d, (hit, tot) in sorted(by_defect.items())},
        "program_length": {
            "min": min(lengths),
            "max": max(lengths),
            "mean": round(sum(lengths) / len(lengths), 1),
        },
    }


def main() -> int:
    m = evaluate()
    if "--json" in sys.argv:
        print(json.dumps(m, indent=2))
    else:
        print(f"surface_cli deterministic eval — {m['total']} cases "
              f"({m['defective']} defective, {m['valid']} valid)")
        print(f"  recall (defects caught):     {m['recall'] * 100:.0f}%")
        print(f"  false-positive rate (valid): {m['false_positive_rate'] * 100:.0f}%")
        print(f"  by defect category:          {m['by_defect']}")
        print(f"  program length (lines):      {m['program_length']}")
        if m["missed"]:
            print(f"  MISSED defects: {m['missed']}")
        if m["false_positives"]:
            print(f"  FALSE POSITIVES: {m['false_positives']}")

    # Regression gate: every seeded defect must be caught (recall 100%) and no valid program may be
    # wrongly flagged — so the helper's quality can't silently regress as the language evolves.
    ok = not m["missed"] and not m["false_positives"]
    print("EVAL OK" if ok else "EVAL FAIL", file=sys.stderr)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
