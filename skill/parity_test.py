#!/usr/bin/env python3
"""
parity_test.py — pin the GE skills' planner verb set to the authoritative intent model.

The TS side (`packages/contracts/src/intent.ts` IntentSchema) is authoritative; the planner
skill's `INTENTS` set must mirror it exactly. This is the lockstep guard called out in
CLAUDE.md ("keep scripts/parse_plan.py ... in lockstep with packages/contracts"). It is
dependency-free (stdlib only) so it runs anywhere the skill bundle does.

Usage:
  python3 skill/parity_test.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "m365-command-planner" / "scripts"))
from parse_plan import CONTEXT_HINTS, INTENTS, parse_plan  # noqa: E402

# The general, Copilot-altitude verbs (docs/EXPERIENCE.md §1 / IntentSchema).
# Scope is an orthogonal axis, never a verb; the deleted task-verbs (regen-clause,
# resolve-comment, draft-slides, synthesize, meeting-notes) are scopes/closures of these.
EXPECTED_INTENTS = {
    "ask",
    "summarize",
    "explain",
    "rewrite",
    "review",
    "visualize",
    "draft",
    "notes",
}

DELETED_INTENTS = {"assist", "regen-clause", "resolve-comment",
                   "draft-slides", "synthesize", "meeting-notes"}
EXPECTED_CONTEXT_HINTS = {
    "incremental",
    "inline-preferred",
    "reference-preferred",
    "upload-preferred",
    "code-execution-preferred",
    "analytical",
    "full-scope",
}


def main() -> int:
    failures = []

    if INTENTS != EXPECTED_INTENTS:
        missing = EXPECTED_INTENTS - INTENTS
        extra = INTENTS - EXPECTED_INTENTS
        if missing:
            failures.append(f"planner INTENTS missing verbs: {sorted(missing)}")
        if extra:
            failures.append(f"planner INTENTS has stray verbs: {sorted(extra)}")

    leaked = INTENTS & DELETED_INTENTS
    if leaked:
        failures.append(f"planner INTENTS still carries deleted task-verbs: {sorted(leaked)}")

    if CONTEXT_HINTS != EXPECTED_CONTEXT_HINTS:
        missing = EXPECTED_CONTEXT_HINTS - CONTEXT_HINTS
        extra = CONTEXT_HINTS - EXPECTED_CONTEXT_HINTS
        if missing:
            failures.append(f"planner CONTEXT_HINTS missing hints: {sorted(missing)}")
        if extra:
            failures.append(f"planner CONTEXT_HINTS has stray hints: {sorted(extra)}")

    sample = """```plan
intent draft
surface excel
scope document
ground this
context analytical
context full-scope
context upload-preferred
context code-execution-preferred
step produce a chart-ready risk table
```"""
    parsed = parse_plan(sample)
    if parsed["errors"]:
        failures.append(f"context sample produced parse errors: {parsed['errors']}")
    if parsed["plan"]["context"] != [
        "analytical",
        "full-scope",
        "upload-preferred",
        "code-execution-preferred",
    ]:
        failures.append(f"context sample parsed wrong hints: {parsed['plan']['context']}")

    unsafe = parse_plan("""```plan
intent ask
surface excel
context run-python-now
step inspect
```""")
    if not any("unknown context hint" in e for e in unsafe["errors"]):
        failures.append("unsafe context hint was not rejected")

    if failures:
        print("PARITY FAIL")
        for f in failures:
            print(f"  - {f}")
        return 1

    print(
        "PARITY OK — planner INTENTS and CONTEXT_HINTS match the typed contract; "
        "context hints parse fail-closed"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
