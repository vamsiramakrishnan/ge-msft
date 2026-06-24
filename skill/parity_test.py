#!/usr/bin/env python3
"""
parity_test.py — pin the GE skills' planner verb set to the authoritative 7-verb model.

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
from parse_plan import INTENTS  # noqa: E402

# The seven general, Copilot-altitude verbs (docs/EXPERIENCE.md §1 / IntentSchema).
# Scope is an orthogonal axis, never a verb; the deleted task-verbs (regen-clause,
# resolve-comment, draft-slides, synthesize, meeting-notes) are scopes/closures of these.
EXPECTED_INTENTS = {"ask", "summarize", "explain", "rewrite", "review", "draft", "notes"}

DELETED_INTENTS = {"assist", "regen-clause", "resolve-comment",
                   "draft-slides", "synthesize", "meeting-notes"}


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

    if failures:
        print("PARITY FAIL")
        for f in failures:
            print(f"  - {f}")
        return 1

    print(f"PARITY OK — planner INTENTS == the 7 verbs {sorted(EXPECTED_INTENTS)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
