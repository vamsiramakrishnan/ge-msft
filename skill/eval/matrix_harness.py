#!/usr/bin/env python3
"""
matrix_harness.py — the DETERMINISTIC three-condition matrix from ADR-0008 §Consequences.

ADR-0008 says the helper ships only if it improves execution quality across three conditions:
base model · Surface Commander · Surface Commander + `surface-cli`. The live form of that
comparison needs a model; this harness replays RECORDED (or synthetic, representative) model
outputs per condition — `eval/matrix/*.jsonl`, one row per task × condition, each with the turn
sequence the model produced — through the same deterministic machinery the runtime and the skill
use (`parse_commands` + `surface_cli.analyze`), with no network and no host.

Per condition it reports the ADR's metrics:
  • first-pass parse success   — the FIRST turn is a complete, error-free `cmd` program;
  • capability-violation rate  — a turn used a verb or `/kind` outside the row's signature;
  • unsupported-verb rate      — a turn used a verb the language does not have;
  • effect-budget rejection    — a turn exceeded the per-turn effect budget (manifest limits);
  • average repair turns       — turns after the first until a valid program (capped at the
                                 runtime's two repair turns; "exhausted" when none was valid);
  • mean program length        — non-blank lines of the final program;
  • helper real-defect share   — for `commander+cli` rows, the share of CLI checks that found a
                                 defect `parse_commands` alone did not (unbound bindings, scope,
                                 type/column laws, recipe arity, budgets).

Each row carries an `expect` block; a mismatch fails the run, and so does the ship bar: the
`commander+cli` condition must not be worse than `commander` on first-pass success or violation
rate. Writes `eval/matrix-report.json`.

Usage:
  python3 eval/matrix_harness.py          # table + gate (exit 1 on a failed expectation / ship bar)
  python3 eval/matrix_harness.py --json
  python3 eval/matrix_harness.py --no-report   # skip writing eval/matrix-report.json
"""

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "m365-surface-commander" / "scripts"))
import parse_commands  # noqa: E402
import surface_cli  # noqa: E402
from surface_cli.budget import DEFAULT_LIMITS  # noqa: E402
from surface_cli.checker import _group_recipes  # noqa: E402

MATRIX_DIR = HERE / "matrix"
# A generated artifact (git-ignored): rewritten on every run, never hand-edited.
REPORT = HERE / "matrix-report.json"
CONDITIONS = ("base", "commander", "commander+cli")
MAX_REPAIR_TURNS = 2  # the runtime's MAX_REPAIR_TURNS (assist-session.ts)

# The per-turn budget is the generated manifest's `limits` block (surface_cli.budget reads it from
# generated_language.LIMITS) — the same numbers the runtime gate re-checks, never a local constant.
_LIMITS = dict(DEFAULT_LIMITS)


def _load_rows():
    rows = []
    for path in sorted(MATRIX_DIR.glob("*.jsonl")):
        for lineno, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if raw.strip():
                row = json.loads(raw)
                row["_where"] = f"{path.name}:{lineno}"
                rows.append(row)
    return rows


def _structural_errors(text):
    """The grammar's verdict on one turn: fence + per-line parse, with `def … end` recipes grouped
    the way the TS `parseProgramBlock` groups them (the Python `parse_block` mirrors the line-level
    `parseCommandBlock`; recipe grouping lives in the CLI's checker, reused here so a recipe call
    is a call, not an unknown verb). Returns (block, errors)."""
    parsed = parse_commands.parse_block(text)
    block = parsed["block"]
    if block is None:
        return None, []
    recipes = {}
    remaining, recipe_errors = _group_recipes(block.splitlines(), recipes)
    if not recipes and not recipe_errors:
        return block, [c["error"] for c in parsed["commands"] if "error" in c]
    regrouped = parse_commands.parse_block("```cmd\n" + "\n".join(remaining) + "\n```")
    errors = list(recipe_errors)
    for c in regrouped["commands"]:
        if "error" not in c:
            continue
        # A call to a recipe defined in this block parses as a call (the line parser alone cannot
        # know the name); any other unknown verb is a real structural error.
        if any(c["error"].startswith(f"unknown verb '{name}'") for name in recipes):
            continue
        errors.append(c["error"])
    return block, errors


def _classify_turn(text, capabilities):
    """One model turn → what the deterministic layers say about it."""
    block, structural_errors = _structural_errors(text)
    analysis = surface_cli.analyze(text, set(capabilities)) if block is not None else None
    # The checker re-reports parse errors; keep only what the preflight ADDS (scope, laws, arity…).
    cli_errors = [e for e in (analysis["errors"] if analysis else []) if e not in structural_errors]
    budget = bool(analysis) and surface_cli._budget_exceeded(analysis, _LIMITS)
    unsupported = any("unknown verb" in e for e in structural_errors)
    violation = any("not in this turn" in e for e in cli_errors)
    # "First-pass parse success" is the grammar's verdict alone; `valid` is what the runtime would
    # accept without a repair turn (grammar + scope/type laws + budget).
    parses = block is not None and not structural_errors
    valid = parses and not cli_errors and not budget
    # A defect only the preflight sees: the structural parser accepted the program, the CLI did not.
    cli_only = parses and bool(cli_errors or budget)
    length = len([ln for ln in (block or "").splitlines() if ln.strip() and not ln.strip().startswith("#")])
    return {
        "has_block": block is not None,
        "parses": parses,
        "valid": valid,
        "unsupported": unsupported,
        "violation": violation,
        "budget": budget,
        "cli_only_defect": cli_only,
        "length": length,
        "errors": structural_errors + cli_errors,
    }


def _run_row(row):
    turns = [_classify_turn(t, row["capabilities"]) for t in row["turns"]]
    first = turns[0]
    repair_turns = None
    for i, t in enumerate(turns):
        if t["valid"]:
            repair_turns = i
            break
    exhausted = repair_turns is None or repair_turns > MAX_REPAIR_TURNS
    final = turns[repair_turns] if repair_turns is not None and not exhausted else turns[-1]
    observed = {
        "first_pass": first["parses"],
        "repair_turns": MAX_REPAIR_TURNS if exhausted else repair_turns,
        "exhausted": exhausted,
        "unsupported": any(t["unsupported"] for t in turns),
        "violation": any(t["violation"] for t in turns),
        "budget": any(t["budget"] for t in turns),
        "cli_found_defect": any(t["cli_only_defect"] for t in turns),
        "program_length": final["length"],
    }
    failures = []
    for key, want in row.get("expect", {}).items():
        if observed.get(key) != want:
            failures.append(f"{key}: expected {want!r}, observed {observed.get(key)!r}")
    return observed, failures, turns


def evaluate():
    rows = _load_rows()
    per = {c: [] for c in CONDITIONS}
    failures = []
    for row in rows:
        observed, row_failures, turns = _run_row(row)
        per[row["condition"]].append(observed)
        for f in row_failures:
            failures.append(f"[{row['task']} / {row['condition']} @ {row['_where']}] {f} — errors: "
                            + "; ".join(e for t in turns for e in t["errors"])[:300])

    def rate(items, key):
        return round(sum(1 for i in items if i[key]) / len(items), 3) if items else None

    summary = {}
    for c in CONDITIONS:
        items = per[c]
        summary[c] = {
            "rows": len(items),
            "first_pass_parse_success": rate(items, "first_pass"),
            "capability_violation_rate": rate(items, "violation"),
            "unsupported_verb_rate": rate(items, "unsupported"),
            "effect_budget_rejection_rate": rate(items, "budget"),
            "repair_exhausted_rate": rate(items, "exhausted"),
            "avg_repair_turns": round(sum(i["repair_turns"] for i in items) / len(items), 2) if items else None,
            "mean_program_length": round(sum(i["program_length"] for i in items) / len(items), 1) if items else None,
        }
    cli_rows = per["commander+cli"]
    summary["commander+cli"]["helper_real_defect_share"] = rate(cli_rows, "cli_found_defect")

    # The ADR's ship bar: the helper must not make things worse than the skill alone.
    ship_bar = []
    a, b = summary["commander"], summary["commander+cli"]
    if a["rows"] and b["rows"]:
        if b["first_pass_parse_success"] < a["first_pass_parse_success"] - 1e-9 and b["avg_repair_turns"] > a["avg_repair_turns"]:
            ship_bar.append("commander+cli has lower first-pass success AND more repair turns than commander")
        if b["capability_violation_rate"] > a["capability_violation_rate"] + 1e-9 and b["repair_exhausted_rate"] > a["repair_exhausted_rate"]:
            ship_bar.append("commander+cli leaves more capability violations unrepaired than commander")
    return {"rows": len(rows), "limits": DEFAULT_LIMITS, "conditions": summary,
            "expectation_failures": failures, "ship_bar_failures": ship_bar}


def _table(report):
    cols = ["first_pass_parse_success", "capability_violation_rate", "unsupported_verb_rate",
            "effect_budget_rejection_rate", "avg_repair_turns", "repair_exhausted_rate",
            "mean_program_length", "helper_real_defect_share"]
    lines = [f"ADR-0008 evaluation matrix — {report['rows']} rows, effect budget {report['limits']['max_effects']}/turn"]
    head = f"{'metric':32}" + "".join(f"{c:>16}" for c in CONDITIONS)
    lines.append(head)
    for col in cols:
        vals = []
        for c in CONDITIONS:
            v = report["conditions"][c].get(col)
            vals.append(f"{'-' if v is None else v:>16}")
        lines.append(f"{col:32}" + "".join(vals))
    return "\n".join(lines)


def main() -> int:
    report = evaluate()
    if "--no-report" not in sys.argv:
        REPORT.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    if "--json" in sys.argv:
        print(json.dumps(report, indent=2))
    else:
        print(_table(report))
        for f in report["expectation_failures"]:
            print("EXPECTATION FAILED", f)
        for f in report["ship_bar_failures"]:
            print("SHIP BAR FAILED", f)
    ok = not report["expectation_failures"] and not report["ship_bar_failures"]
    print("MATRIX OK" if ok else "MATRIX FAIL", file=sys.stderr)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
