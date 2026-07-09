#!/usr/bin/env python3
"""
surface_cli.py — the runnable CLI entry for the deterministic PREFLIGHT compiler (ADR-0008 §4).

The implementation lives in the `surface_cli/` package (types · generated_language · parser · checker
· normalizer · budget); this file is the thin entry that wires argparse to it and renders the report.
It answers, for one `cmd` program: does it parse? does it use only this turn's capabilities? what are
the bindings and their inferred types? how many reads/effects/cells vs budget? which effects depend on
which (the approval-preview groups)?

It is a PURE compiler tool: it NEVER calls Office.js or Graph, acquires tokens, discovers
capabilities, executes code, or mutates anything. It reuses the manifest-wired parser so it can never
disagree with the runtime on the verb set.

Usage:
  surface_cli.py check     --surface excel [--capabilities a,b,c] < program
  surface_cli.py budget    --max-effects 8 --max-reads 8 --max-cells 10000 < program
  surface_cli.py plan      < program
  surface_cli.py normalize < program     # reorder into OBSERVE -> DERIVE -> EFFECT canonical form
  surface_cli.py help shape               # generated targeted command playbook
  surface_cli.py --self-test
"""

import argparse
import json
import sys
from pathlib import Path

# scripts/ on the path so `import surface_cli` resolves to the PACKAGE (it shadows this entry file for
# import; this file only runs as __main__), and so the package can reach parse_commands.py.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from surface_cli import analyze, normalize, _budget_exceeded, _cell_count  # noqa: E402
from command_help import render as render_command_help  # noqa: E402


# ───────────────────────────── rendering ─────────────────────────────


def _render(result, limits, show_budget=True, show_plan=True):
    out = []
    valid = not result["errors"]
    out.append(("VALID " if valid else "INVALID ") + result["version"])
    for err in result["errors"]:
        out.append(f"  error: {err}")

    if result["bindings"]:
        out.append("Bindings")
        for b in result["bindings"]:
            out.append(f"  ${b['name']} : {b['type']}")

    if result["effects"]:
        out.append("Effects")
        for i, e in enumerate(result["effects"]):
            tgt = e["target"] or "(anchored)"
            dep = next((f"  depends-on e{j + 1}" for (a, j) in result["deps"] if a == i), "")
            out.append(f"  e{i + 1} {e['verb']:<7} {tgt}{dep}")

    if show_plan and result["deps"]:
        out.append("Plan (dependency groups)")
        for (i, j) in result["deps"]:
            out.append(
                f"  e{i + 1} ({result['effects'][i]['verb']}) ← e{j + 1} ({result['effects'][j]['verb']})"
            )

    if show_budget:
        n_reads = len(result["reads"])
        n_eff = len(result["effects"])
        n_cells = sum(_cell_count(e["range"]) for e in result["effects"])
        out.append("Budget")
        out.append(f"  reads:   {n_reads} / {limits['max_reads']}")
        out.append(f"  effects: {n_eff} / {limits['max_effects']}")
        out.append(f"  cells:   {n_cells} / {limits['max_cells']}")
        if n_eff > limits["max_effects"]:
            out.append(f"  BUDGET EXCEEDED: {n_eff} effects > {limits['max_effects']}")
        if n_reads > limits["max_reads"]:
            out.append(f"  BUDGET EXCEEDED: {n_reads} reads > {limits['max_reads']}")
        if n_cells > limits["max_cells"]:
            out.append(f"  BUDGET EXCEEDED: {n_cells} cells > {limits['max_cells']}")

    ext = [f"e{i + 1}" for i, e in enumerate(result["effects"]) if e["external"]]
    out.append("Risk")
    out.append(f"  external effects: {', '.join(ext) if ext else 'none'}")
    out.append("  irreversible effects: none")
    return "\n".join(out)


# ───────────────────────────── CLI ─────────────────────────────


def _run(argv):
    ap = argparse.ArgumentParser(prog="surface_cli", description="m365-cli preflight compiler")
    ap.add_argument(
        "command",
        choices=["check", "budget", "plan", "explain", "normalize", "help"],
        nargs="?",
    )
    ap.add_argument("topic", nargs="?")
    ap.add_argument("--surface")
    ap.add_argument("--capabilities", help="comma-separated verbs live this turn (optional scope)")
    ap.add_argument("--max-effects", type=int, default=8)
    ap.add_argument("--max-reads", type=int, default=8)
    ap.add_argument("--max-cells", type=int, default=10000)
    ap.add_argument("--json", action="store_true", help="machine-readable output (for tests/orchestration)")
    ap.add_argument("--self-test", action="store_true")
    args = ap.parse_args(argv)

    if args.self_test:
        return _self_test()

    if args.command == "help":
        print(render_command_help(args.topic))
        return 0

    caps = set(args.capabilities.split(",")) if args.capabilities else None
    limits = {"max_effects": args.max_effects, "max_reads": args.max_reads, "max_cells": args.max_cells}
    source = sys.stdin.read()

    # `normalize` reorders the program into canonical form; it does not analyze/budget.
    if args.command == "normalize":
        lines, notes = normalize(source)
        for line in lines:
            print(line)
        for note in notes:
            print(f"# note: {note}", file=sys.stderr)
        return 0

    result = analyze(source, caps)

    if args.json:
        print(json.dumps({**result, "budgetExceeded": _budget_exceeded(result, limits)}, indent=2))
    else:
        cmd = args.command or "check"
        print(
            _render(
                result,
                limits,
                show_budget=cmd in ("budget", "check", "explain"),
                show_plan=cmd in ("plan", "check", "explain"),
            )
        )

    # Exit non-zero on a real defect so the model/orchestrator can branch on it.
    return 1 if (result["errors"] or _budget_exceeded(result, limits)) else 0


# ───────────────────────────── self-test ─────────────────────────────


def _self_test() -> int:
    failures = []

    # 1. The canonical top-N report: pure pipeline + spill + table/chart over the spilled range.
    prog = """```cmd
let $top = read Sales!A1:D200 | sort ARR desc | head 10
spill Report!A1 = ($top)
table Report!A1:D11 name=TopAccounts
chart column Report!A1:D11 title="Top Accounts"
done
```"""
    r = analyze(prog)
    if r["errors"]:
        failures.append(f"clean top-N program had errors: {r['errors']}")
    if not any(b["name"] == "top" and b["type"] == "Table" for b in r["bindings"]):
        failures.append("did not bind $top : Table")
    dep_pairs = set(r["deps"])
    eff_verbs = [e["verb"] for e in r["effects"]]
    spill_i = eff_verbs.index("spill")
    if not all((eff_verbs.index(v), spill_i) in dep_pairs for v in ("table", "chart")):
        failures.append(f"table/chart did not depend on spill; deps={r['deps']} verbs={eff_verbs}")

    # 2. Capability scope: a verb absent from the turn's signature is rejected.
    r2 = analyze("```cmd\nchart column A1:B2\ndone\n```", capabilities={"set", "table"})
    if not any("not in this turn" in e for e in r2["errors"]):
        failures.append("capability scope did not reject an unavailable verb")

    # 3. Unbound $var in spill is caught.
    r3 = analyze("```cmd\nspill Report!A1 = ($ghost)\ndone\n```")
    if not any("unbound" in e for e in r3["errors"]):
        failures.append("unbound spill var not caught")

    # 4. Budget: a flat 10x10 set of effects exceeds max-effects=8.
    many = "```cmd\n" + "\n".join(f"set A{i} 1" for i in range(1, 11)) + "\ndone\n```"
    rb = analyze(many)
    if not _budget_exceeded(rb, {"max_effects": 8, "max_reads": 8, "max_cells": 10000}):
        failures.append("10 effects did not exceed max-effects=8")

    # 5. External-effect risk surfaced.
    rext = analyze('```cmd\nmail "hi"\ndone\n```')
    if not rext["effects"] or not rext["effects"][0]["external"]:
        failures.append("mail not flagged as an external effect")

    # 6. A parse error makes the program INVALID.
    rerr = analyze("```cmd\nset A1\ndone\n```")
    if not rerr["errors"]:
        failures.append("malformed `set A1` did not error")

    # 7. The /<kind> specialized surface: scoped by KIND, flagged external when it reaches the estate.
    rinv = analyze(
        '```cmd\n/insert-image base64=AAA alt="chart"\n/post-channel-message text="hi"\ndone\n```',
        capabilities={"insert-image"},
    )
    if not any("post-channel-message" in e for e in rinv["errors"]):
        failures.append("/ surface did not scope an unavailable kind")
    if not any(e["external"] for e in rinv["effects"]):
        failures.append("/post-channel-message not flagged external")
    if len([e for e in rinv["effects"] if e["verb"].startswith("/")]) != 2:
        failures.append("/ invokes not counted as effects")

    # 8. An unknown /<kind> is a parse error (did-you-mean from parse_commands).
    rbad = analyze("```cmd\n/insert-imag base64=AAA\ndone\n```")
    if not any("unknown capability" in e for e in rbad["errors"]):
        failures.append("unknown /kind not rejected")

    # 9. normalize hoists DERIVE above EFFECT, keeps `done` last (OBSERVE→DERIVE→EFFECT→CONTROL).
    nlines, _ = normalize(
        "```cmd\nspill Report!A1 = ($top)\nlet $top = read Sales!A1:D9 | head 10\ndone\n```"
    )
    if nlines != ["let $top = read Sales!A1:D9 | head 10", "spill Report!A1 = ($top)", "done"]:
        failures.append(f"normalize did not reorder to OBSERVE→DERIVE→EFFECT: {nlines}")

    # 10. a read AFTER an effect is flagged as a fresh-observation (VERIFY) boundary.
    _, nnotes = normalize("```cmd\nset A1 5\nread B1\ndone\n```")
    if not any("VERIFY" in n for n in nnotes):
        failures.append("normalize did not flag read-after-effect as a VERIFY boundary")

    # 11. context strategy is an OBSERVE/read phase, not an effect; unknown hints fail closed.
    rctx = analyze(
        "```cmd\ncontext analytical full-scope upload-preferred code-execution-preferred\ndone\n```"
    )
    if rctx["errors"]:
        failures.append(f"context strategy program had errors: {rctx['errors']}")
    if rctx["reads"] != ["context analytical full-scope upload-preferred code-execution-preferred"]:
        failures.append(f"context strategy was not counted as one read: {rctx['reads']}")
    if rctx["effects"]:
        failures.append(f"context strategy was incorrectly counted as an effect: {rctx['effects']}")
    rctx_bad = analyze("```cmd\ncontext run-python-now\ndone\n```")
    if not any("unknown context hint" in e for e in rctx_bad["errors"]):
        failures.append("unknown context hint not rejected")

    # 12. Workspace artifacts are local observe/workbench steps, not host effects.
    rws = analyze(
        "```cmd\nsave schedule.tsv = read 'Daily schedule'!B3:I53\ngrep schedule.tsv \"Deep Work\" context=1\ndone\n```"
    )
    if rws["errors"]:
        failures.append(f"workspace program had errors: {rws['errors']}")
    if len(rws["reads"]) != 2 or rws["effects"]:
        failures.append(f"workspace commands were misclassified: reads={rws['reads']} effects={rws['effects']}")

    # 13. Generated targeted help is available through the CLI front door.
    help_text = render_command_help("shape")
    if "Command: shape" not in help_text or "pp:shape:slideId:shapeId" not in help_text:
        failures.append(f"targeted shape help did not render from generated manifest: {help_text}")

    if failures:
        print("SURFACE-CLI SELF-TEST FAIL", file=sys.stderr)
        for f in failures:
            print(f"  - {f}", file=sys.stderr)
        return 1
    print(
        "SURFACE-CLI SELF-TEST OK — check/budget/plan/normalize, capability scope, dep inference, risk"
    )
    return 0


if __name__ == "__main__":
    sys.exit(_run(sys.argv[1:]))
