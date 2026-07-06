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
    """Extract the verb set from capability-map.md: the read columns,
    every `\\`verb\\`` named in the Writes table's first column, plus the control verbs."""
    cap_map = HERE / "m365-surface-commander" / "references" / "capability-map.md"
    text = cap_map.read_text(encoding="utf-8")
    verbs = set(DOCUMENTED_CONTROL_VERBS)
    # Read verbs are referenced inline in the read table and read-body prose.
    for v in sorted(parse_commands.READ_VERBS):
        if f"`{v}`" in text:
            verbs.add(v)
    # Write verbs: rows of the Writes table begin with `| \`<verb>\``.
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("| `") and "|" in stripped[1:]:
            cell = stripped.split("|", 2)[1].strip()  # first column
            if cell.startswith("`") and cell.endswith("`"):
                verb = cell.strip("`")
                # Slash-specialized capabilities are parsed through the `invoke` arm, not as
                # top-level core verbs, so they are covered by golden invoke cases instead.
                if not verb.startswith("/"):
                    verbs.add(verb)
    return verbs


def _check_manifest_parity() -> list:
    """ADR-0008 §4: the generated language manifest is the SINGLE SOURCE. Assert it agrees with
    BOTH the Python parser (which loads from it) and the capability-map doc — so the TS grammar →
    manifest → parser/doc chain has no drift at any hop. Skipped (with a note) if the manifest is
    absent (a stripped sandbox running on the hardcoded fallback)."""
    failures = []
    manifest_path = parse_commands._MANIFEST_PATH
    if not manifest_path.exists():
        print("  [manifest] note: bundled manifest absent — parser on hardcoded fallback")
        return failures
    data = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest_all = set(data["verbs"]["read"]) | set(data["verbs"]["control"]) | set(data["verbs"]["write"])
    manifest_write = set(data["verbs"]["write"])

    # 1. parser ≡ manifest (guards the loader + that the fallback never silently diverges).
    if set(ALL_VERBS) != manifest_all:
        failures.append(f"  [manifest] parser ALL_VERBS != manifest verbs: "
                        f"parser-only {sorted(set(ALL_VERBS) - manifest_all)}, "
                        f"manifest-only {sorted(manifest_all - set(ALL_VERBS))}")
    # 2. every manifest write verb has a parse arm (drift gate, the other direction).
    if parse_commands.HANDLED_WRITE_VERBS != manifest_write:
        failures.append(f"  [manifest] write verbs without a parse arm: "
                        f"{sorted(manifest_write - parse_commands.HANDLED_WRITE_VERBS)}")
    # 3. doc ≡ manifest (the human capability-map matches the generated source).
    doc_write = _documented_verbs() - DOCUMENTED_CONTROL_VERBS - set(parse_commands.READ_VERBS)
    if doc_write != manifest_write:
        failures.append(f"  [manifest] capability-map writes != manifest writes: "
                        f"doc-only {sorted(doc_write - manifest_write)}, "
                        f"manifest-only {sorted(manifest_write - doc_write)}")
    return failures


def main() -> int:
    failures = []
    failures += _run_corpus()
    failures += _check_done_with_error_invariant()
    failures += _check_verb_set_parity()
    failures += _check_manifest_parity()

    if failures:
        print("PARITY CORPUS FAIL")
        for f in failures:
            print(f)
        return 1

    n = len(_load_cases())
    print(f"PARITY CORPUS OK — {n} golden cases, fail-closed `done`, verb-set + manifest parity")
    return 0


if __name__ == "__main__":
    sys.exit(main())
