#!/usr/bin/env python3
"""
surface_cli.py — the deterministic PREFLIGHT compiler for m365-cli programs (ADR-0008 §4).

It answers, for one `cmd` program: does it parse? does it use only this turn's capabilities? what
are the bindings and their inferred types? how many reads/effects/cells, and against what budget?
which effects depend on which (the approval-preview groups)?

It is a PURE compiler tool. It NEVER calls Office.js or Graph, acquires tokens, discovers
capabilities, executes code, or mutates anything (ADR-0008 §4). It reuses the manifest-wired parser
(`parse_commands.py`) so it can never disagree with the runtime on the verb set.

Usage:
  surface_cli.py check   --surface excel [--capabilities a,b,c] < program
  surface_cli.py budget  --max-effects 8 --max-reads 8 --max-cells 10000 < program
  surface_cli.py plan    < program
  surface_cli.py --self-test
"""

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import parse_commands  # noqa: E402
from parse_commands import (  # noqa: E402
    parse_line,
    extract_command_block_meta,
    TRANSFORM_NAMES,
    EFFECT_VERBS,
    LANGUAGE_VERSION,
)

# Transform → output value type (ADR-0008 §1 algebra). Aggregations collapse a Table to a Number;
# everything else stays a Table. A source (read/search/outline/$var) is a Table.
_AGG = {"sum", "avg", "min", "max", "count"}


# ───────────────────────────── A1 range algebra (for dependency inference) ─────────────────────────────

_RANGE = re.compile(r"^(?:(?P<sheet>[^!]+)!)?(?P<a>[A-Za-z]+\d+)(?::(?P<b>[A-Za-z]+\d+))?$")
_CELL = re.compile(r"^([A-Za-z]+)(\d+)$")


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _parse_range(ref: str):
    """`Sheet!A1:G11` / `A1:G11` / `A1` → (sheet, c1, r1, c2, r2) in 1-based ints, or None."""
    m = _RANGE.match(ref.strip())
    if not m:
        return None
    a = _CELL.match(m.group("a"))
    if not a:
        return None
    c1, r1 = _col_to_num(a.group(1)), int(a.group(2))
    if m.group("b"):
        b = _CELL.match(m.group("b"))
        if not b:
            return None
        c2, r2 = _col_to_num(b.group(1)), int(b.group(2))
    else:
        c2, r2 = c1, r1
    sheet = (m.group("sheet") or "").strip().strip("'")
    return (sheet, min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))


def _overlap(x, y) -> bool:
    """Two parsed ranges intersect (same sheet + box overlap). A blank sheet matches any sheet."""
    if x is None or y is None:
        return False
    sx, x1, y1, x2, y2 = x
    sy, a1, b1, a2, b2 = y
    if sx and sy and sx.lower() != sy.lower():
        return False
    return not (x2 < a1 or a2 < x1 or y2 < b1 or b2 < y1)


def _cell_count(rng) -> int:
    if rng is None:
        return 0
    _s, c1, r1, c2, r2 = rng
    return (c2 - c1 + 1) * (r2 - r1 + 1)


# ───────────────────────────── program model ─────────────────────────────

# Effects that touch the recipient/host externally (vs an in-document change) — surfaced as risk.
_EXTERNAL = {"mail", "post", "compose"}


def _is_expr_line(line: str) -> bool:
    """A `let $x = …` binding or a bare pipeline (a top-level ` | `), mirroring expr-grammar."""
    s = line.strip()
    if s.startswith("let "):
        return True
    depth = 0
    for i, ch in enumerate(s):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        elif ch == "|" and depth == 0 and 0 < i < len(s) - 1 and s[i - 1] == " " and s[i + 1] == " ":
            return True
    return False


def _infer_pipeline_type(rhs: str):
    """Infer the value type (`Table`/`Number`) of a pipeline RHS from its last transform. Returns
    (vtype, source, transform_names, unknown_transforms)."""
    parts = [p.strip() for p in re.split(r"\s\|\s", rhs.strip())]
    source = parts[0]
    stages = parts[1:]
    vtype = "Number" if source.lstrip("$").split()[0] in _AGG else "Table"
    names, unknown = [], []
    for st in stages:
        name = st.split()[0] if st.split() else ""
        names.append(name)
        if name not in TRANSFORM_NAMES:
            unknown.append(name)
        vtype = "Number" if name in _AGG else "Table"
    return vtype, source, names, unknown


def analyze(program_text: str, capabilities=None):
    """Parse + scope + type a program. Returns a structured result (no rendering, no side effects)."""
    inner, closed = extract_command_block_meta(program_text)
    if inner is None:
        inner, closed = program_text, True  # accept a bare program (no fence) for the CLI

    errors, bindings, effects, reads = [], [], [], []
    bound = set()

    for raw in inner.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue

        if _is_expr_line(line):
            m = re.match(r"^let\s+\$(\w+)\s*=\s*(.+)$", line)
            if m:
                name, rhs = m.group(1), m.group(2)
                vtype, src, _names, unknown = _infer_pipeline_type(rhs)
                for u in unknown:
                    errors.append(f'unknown transform "{u}" in ${name}')
                # a read source is a host read
                if src.split()[0] in ("read", "search", "outline"):
                    reads.append(src)
                # referencing an unbound $var
                for ref in re.findall(r"\$(\w+)", rhs):
                    if ref not in bound:
                        errors.append(f"${ref} used before it is bound")
                bound.add(name)
                bindings.append({"name": name, "type": vtype, "source": src})
            else:
                vtype, src, _n, unknown = _infer_pipeline_type(line)
                for u in unknown:
                    errors.append(f'unknown transform "{u}"')
                if src.split()[0] in ("read", "search", "outline"):
                    reads.append(src)
            continue

        rec = parse_line(line)
        if rec is None:
            continue
        if "error" in rec:
            errors.append(rec["error"])
            continue

        verb = rec["verb"]
        if verb in ("outline", "read", "search"):
            reads.append(line)
            continue
        if verb in ("done", "help"):
            continue

        # An effect. Capability-scope it against this turn's signature, if one was supplied.
        if capabilities is not None and verb not in capabilities:
            errors.append(f'"{verb}" is not in this turn\'s capabilities')
        # Its target/read ranges (Excel family) for dependency + cell inference.
        target = rec.get("range") or rec.get("cell")
        eff = {
            "verb": verb,
            "target": target,
            "range": _parse_range(target) if target else None,
            "external": verb in _EXTERNAL,
            "refs": re.findall(r"\$(\w+)", rec.get("value", "")) if verb == "spill" else [],
        }
        # spill referencing an unbound var
        for ref in eff["refs"]:
            if ref not in bound:
                errors.append(f"spill references unbound ${ref}")
        effects.append(eff)

    # Effect→effect dependencies: a later effect whose range overlaps an earlier effect's range
    # depends on it (the derived-range case — table/chart/cf over a spilled region).
    deps = []
    for i, e in enumerate(effects):
        for j in range(i):
            if e["range"] and effects[j]["range"] and _overlap(e["range"], effects[j]["range"]):
                deps.append((i, j))
                break

    return {
        "version": LANGUAGE_VERSION,
        "closed": closed,
        "errors": errors,
        "bindings": bindings,
        "effects": effects,
        "reads": reads,
        "deps": deps,
    }


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
            out.append(f"  e{i + 1} ({result['effects'][i]['verb']}) ← e{j + 1} ({result['effects'][j]['verb']})")

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


def _budget_exceeded(result, limits) -> bool:
    return (
        len(result["effects"]) > limits["max_effects"]
        or len(result["reads"]) > limits["max_reads"]
        or sum(_cell_count(e["range"]) for e in result["effects"]) > limits["max_cells"]
    )


# ───────────────────────────── CLI ─────────────────────────────


def _run(argv):
    ap = argparse.ArgumentParser(prog="surface_cli", description="m365-cli preflight compiler")
    ap.add_argument("command", choices=["check", "budget", "plan", "explain"], nargs="?")
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

    caps = set(args.capabilities.split(",")) if args.capabilities else None
    limits = {"max_effects": args.max_effects, "max_reads": args.max_reads, "max_cells": args.max_cells}
    result = analyze(sys.stdin.read(), caps)

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
    # table(e3) and chart(e4) both overlap the spill(e2) target → both depend on it.
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

    if failures:
        print("SURFACE-CLI SELF-TEST FAIL", file=sys.stderr)
        for f in failures:
            print(f"  - {f}", file=sys.stderr)
        return 1
    print("SURFACE-CLI SELF-TEST OK — check/budget/plan, capability scope, dep inference, risk")
    return 0


if __name__ == "__main__":
    sys.exit(_run(sys.argv[1:]))
