#!/usr/bin/env python3
"""
parity_corpus_test.py — drive every case in `golden-corpus.jsonl` through the Python command
checker (`parse_commands`) and assert its structured verb/error output matches the expectation.

This is the parity guard called out in review Finding #8/#10: the TS grammar
(`packages/contracts/src/command-grammar.ts`) is authoritative and the Python checker
(`m365-surface-commander/scripts/parse_commands.py`) is hand-maintained alongside it with no
shared test. The golden corpus pins the behaviors that BOTH parsers must agree on — each flat
command, quoted/escaped values, an unknown verb, an unclosed fence, and (the fail-closed
invariant) a block that contains BOTH `done` AND a parse error.

Critically it asserts the actuation gate's safety property: **a block with any parse error is
NOT reported complete even if it also contains `done`** (README §"What we learned": the add-in
must "not honor `done` if the same block had parse errors").

Dependency-free (stdlib only) so it runs anywhere the skill bundle does.

Usage:
  python3 skill/parity_corpus_test.py
"""

import json
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE / "m365-surface-commander" / "scripts"))

import parse_commands  # noqa: E402
from parse_commands import (  # noqa: E402
    ALL_VERBS,
    block_is_complete,
    extract_command_block,
    parse_block,
    parse_line,
)

CORPUS = HERE / "golden-corpus.jsonl"


def _load_cases():
    cases = []
    with CORPUS.open(encoding="utf-8") as fh:
        for lineno, raw in enumerate(fh, 1):
            raw = raw.strip()
            if not raw:
                continue
            try:
                cases.append((lineno, json.loads(raw)))
            except json.JSONDecodeError as exc:
                raise SystemExit(f"golden-corpus.jsonl:{lineno}: invalid JSON — {exc}")
    return cases


_BLOCK_KEYS = ("complete", "no_block", "has_error")


def _is_block(text: str, expect: dict) -> bool:
    """A corpus case exercises the block path iff it opens a ```cmd fence OR asserts a block-level
    property (`complete` / `no_block` / `has_error`) — the latter covers a prose-only `no_block`
    case whose input deliberately has no fence."""
    return "```cmd" in text or any(k in expect for k in _BLOCK_KEYS)


def _check_flat(expect: dict) -> list:
    """A single-line case is exercised through parse_line (the per-line contract)."""
    errs = []
    rec = parse_line(expect["_input"])

    if "error_substring" in expect:
        if rec is None or "error" not in rec:
            errs.append(f"expected an error, got {rec!r}")
        elif expect["error_substring"] not in rec["error"]:
            errs.append(f"error {rec['error']!r} lacks {expect['error_substring']!r}")
        return errs

    verbs = expect.get("verbs", [])
    if not verbs:
        # A skipped line (blank / comment) → parse_line returns None.
        if rec is not None:
            errs.append(f"expected a skipped line (None), got {rec!r}")
        return errs

    if rec is None or "error" in rec:
        errs.append(f"expected verb {verbs[0]!r}, got {rec!r}")
    elif rec.get("verb") != verbs[0]:
        errs.append(f"expected verb {verbs[0]!r}, got {rec.get('verb')!r}")
    return errs


def _check_block(expect: dict) -> list:
    errs = []
    parsed = parse_block(expect["_input"])
    cmds = parsed["commands"]

    if expect.get("no_block"):
        if parsed["block"] is not None:
            errs.append(f"expected no ```cmd fence, got block {parsed['block']!r}")
        return errs

    got_verbs = [c["verb"] for c in cmds if "verb" in c]
    has_error = any("error" in c for c in cmds)

    if "verbs" in expect and got_verbs != expect["verbs"]:
        errs.append(f"verbs {got_verbs} != expected {expect['verbs']}")
    if "has_error" in expect and has_error != expect["has_error"]:
        errs.append(f"has_error={has_error} != expected {expect['has_error']}")

    if "complete" in expect:
        complete = block_is_complete(parsed)
        if complete != expect["complete"]:
            errs.append(f"complete={complete} != expected {expect['complete']}")
        # Fail-closed invariant: a block with any parse error is never complete.
        if has_error and complete:
            errs.append("block with a parse error was reported complete (fail-open!)")
    return errs


def _run_corpus() -> list:
    failures = []
    for lineno, case in _load_cases():
        expect = dict(case["expect"])
        expect["_input"] = case["input"]
        name = case.get("name", f"line {lineno}")
        check = _check_block if _is_block(case["input"], expect) else _check_flat
        for err in check(expect):
            failures.append(f"  [{name}] {err}")
    return failures


def _check_done_with_error_invariant() -> list:
    """Belt-and-suspenders: a fenced block carrying `done` AND a malformed line must NOT be
    reported complete — the actuation gate stays fail-closed on a partially-broken block."""
    parsed = parse_block("```cmd\nwrit-cells A1 5\ndone\n```")
    has_done = any(c.get("verb") == "done" for c in parsed["commands"])
    has_error = any("error" in c for c in parsed["commands"])
    if not (has_done and has_error):
        return ["  [invariant] expected the block to contain BOTH `done` and a parse error"]
    if block_is_complete(parsed):
        return ["  [invariant] block with `done` + parse error was reported complete (fail-open!)"]
    return []


def _check_verb_set_parity() -> list:
    """Set-equality: parse_commands.ALL_VERBS == the verb set documented in capability-map.md."""
    documented = _documented_verbs()
    py_verbs = set(ALL_VERBS)
    failures = []
    missing = documented - py_verbs
    extra = py_verbs - documented
    if missing:
        failures.append(f"  [verb-set] parse_commands missing documented verbs: {sorted(missing)}")
    if extra:
        failures.append(f"  [verb-set] parse_commands has verbs not in capability-map: {sorted(extra)}")
    return failures


# Control verbs are documented in prose (the actuation/turn loop), not the capability table.
DOCUMENTED_CONTROL_VERBS = {"done", "help"}


def _documented_verbs() -> set:
    """Extract the verb set from capability-map.md: the `outline`/`read`/`search` read columns,
    every `\\`verb\\`` named in the Writes table's first column, plus the control verbs."""
    cap_map = HERE / "m365-surface-commander" / "references" / "capability-map.md"
    text = cap_map.read_text(encoding="utf-8")
    verbs = set(DOCUMENTED_CONTROL_VERBS)
    # Read verbs are referenced inline as `outline` / `read` / `search`.
    for v in ("outline", "read", "search"):
        if f"`{v}`" in text:
            verbs.add(v)
    # Write verbs: rows of the Writes table begin with `| \`<verb>\``.
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("| `") and "|" in stripped[1:]:
            cell = stripped.split("|", 2)[1].strip()  # first column
            if cell.startswith("`") and cell.endswith("`"):
                verbs.add(cell.strip("`"))
    return verbs


def main() -> int:
    failures = []
    failures += _run_corpus()
    failures += _check_done_with_error_invariant()
    failures += _check_verb_set_parity()

    if failures:
        print("PARITY CORPUS FAIL")
        for f in failures:
            print(f)
        return 1

    n = len(_load_cases())
    print(f"PARITY CORPUS OK — {n} golden cases, fail-closed `done` invariant, verb-set equality")
    return 0


if __name__ == "__main__":
    sys.exit(main())
