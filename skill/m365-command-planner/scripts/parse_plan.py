#!/usr/bin/env python3
"""
parse_plan.py — extract the ```plan fenced block from a planner reply and parse its keyword
lines into a structured CommandPlan, emitting short corrective errors for malformed lines.

Sibling of m365-surface-commander/scripts/parse_commands.py. The Office add-in applies the
authoritative parse when it dispatches the plan; this is a lightweight, dependency-free checker.

Usage:
  echo '```plan
  intent review
  surface word
  step rewrite the SLA to 99.9% as a tracked change
  ```' | python3 parse_plan.py
  python3 parse_plan.py --self-test
"""

import json
import re
import sys
from pathlib import Path
from language_manifest import load_manifest, ManifestError

try:
    _LANGUAGE = load_manifest(Path(__file__).with_name("m365-plan-1.0.json"), "m365-plan/1.0")
except ManifestError as error:
    raise SystemExit(str(error)) from error
SCALAR_KEYS = set(_LANGUAGE["scalarKeys"])
LIST_KEYS = set(_LANGUAGE["listKeys"])
BRACKETS = set(_LANGUAGE["brackets"])
ALL_KEYS = SCALAR_KEYS | LIST_KEYS | BRACKETS
INTENTS = set(_LANGUAGE["intents"])
SURFACES = set(_LANGUAGE["surfaces"])
CONFIDENCE = set(_LANGUAGE["confidence"])
CONTEXT_HINTS = set(_LANGUAGE["contextHints"])

_FENCE = re.compile(r"```plan[^\S\n]*\r?\n([\s\S]*?)```", re.IGNORECASE)
_FENCE_OPEN = re.compile(r"```plan[^\S\n]*\r?\n([\s\S]*)$", re.IGNORECASE)


def extract_plan_block(text: str):
    """Return the inner text of the first ```plan fence, or None (→ re-prompt, not an error).
    Tolerates an unclosed fence (a frequent real-world failure mode)."""
    m = _FENCE.search(text)
    if m:
        return m.group(1).strip()
    m = _FENCE_OPEN.search(text)
    if m:
        return re.sub(r"```\s*$", "", m.group(1)).strip()
    return None


def _did_you_mean(key: str):
    import difflib
    near = difflib.get_close_matches(key, sorted(ALL_KEYS - BRACKETS), n=1)
    return f" — did you mean '{near[0]}'?" if near else ""


def parse_line(line: str):
    """Parse one keyword line → ('key', value) | ('bracket', None) | {'error': ...} | None."""
    line = line.strip()
    if not line or line.startswith("#"):
        return None
    parts = line.split(None, 1)
    key = parts[0].lower()
    rest = parts[1].strip() if len(parts) > 1 else ""

    if key in BRACKETS:
        return ("bracket", None)
    if key not in ALL_KEYS:
        return {"error": f"unknown plan keyword '{key}'{_did_you_mean(key)}"}
    if not rest:
        return {"error": f"'{key}' needs a value"}

    if key == "intent" and rest not in INTENTS:
        return {"error": f"unknown intent '{rest}' — expected one of {sorted(INTENTS)}"}
    if key == "surface" and rest.lower() not in SURFACES:
        return {"error": f"unknown surface '{rest}' — expected one of {sorted(SURFACES)}"}
    if key == "confidence" and rest.lower() not in CONFIDENCE:
        return {"error": f"confidence must be high|medium|low — got '{rest}'"}
    if key == "context" and rest.lower() not in CONTEXT_HINTS:
        return {"error": f"unknown context hint '{rest}' — expected one of {sorted(CONTEXT_HINTS)}"}
    if key == "ground":
        rest = rest.strip().strip('"')
    return (key, rest)


def parse_plan(model_text: str):
    """Return {'block', 'plan', 'errors'}. `plan` is the structured CommandPlan."""
    inner = extract_plan_block(model_text)
    if inner is None:
        return {"block": None, "plan": None, "errors": [],
                "note": "no ```plan fence (re-prompt, not an error)"}

    plan = {**{key: None for key in SCALAR_KEYS}, **{key: [] for key in LIST_KEYS}}
    errors = []
    for raw in inner.splitlines():
        rec = parse_line(raw)
        if rec is None:
            continue
        if isinstance(rec, dict):
            errors.append(rec["error"])
            continue
        key, val = rec
        if key == "bracket":
            continue
        if key in SCALAR_KEYS:
            plan[key] = val.lower() if key in ("surface", "confidence") else val
        elif key == "context":
            plan[key].append(val.lower())
        else:
            plan[key].append(val)

    # structural validation (mirrors the grammar's "required" column)
    if not plan["intent"]:
        errors.append("plan is missing 'intent'")
    if not plan["surface"]:
        errors.append("plan is missing 'surface'")
    if not plan["step"] and not plan["clarify"]:
        errors.append("plan needs at least one 'step' (or a 'clarify' to ask first)")
    plan["needs_clarification"] = bool(plan["clarify"])
    return {"block": inner, "plan": plan, "errors": errors}


def _self_test():
    sample = '''**thought** mapping the request to a plan
```plan
intent   rewrite
surface  word
scope    section §4-6
ground   "Vendor Risk Policy v4"
context  incremental
step     rewrite the SLA availability figure to 99.9% as a tracked change
exclude  the indemnity clause — leave unchanged
confidence high
prioritise nonsense
```'''
    result = parse_plan(sample)
    failures = []
    plan = result["plan"]
    if plan["intent"] != "rewrite":
        failures.append("intent did not parse")
    if plan["surface"] != "word":
        failures.append("surface did not parse")
    if plan["context"] != ["incremental"]:
        failures.append(f"context hints did not parse: {plan['context']}")
    if not any("unknown plan keyword 'prioritise'" in e for e in result["errors"]):
        failures.append("unknown keyword was not reported")

    unsafe = parse_plan("""```plan
intent ask
surface excel
context run-python-now
step inspect
```""")
    if not any("unknown context hint" in e for e in unsafe["errors"]):
        failures.append("unsafe context hint was not rejected")

    cross = parse_plan("""```plan
intent draft
surface excel
context analytical
step Prepare a slide-ready summary with source refs and constraints in Excel.
step Resume deck creation only after the user opens PowerPoint and reviews the handoff.
exclude Do not mutate PowerPoint from Excel.
confidence high
```""")
    if cross["errors"]:
        failures.append(f"Supported cross-surface intentions errored: {cross['errors']}")
    for keyword in ("workflow", "source", "target", "phase", "handoff"):
        unsupported = parse_plan(f"```plan\nintent draft\nsurface excel\n{keyword} unsupported\nstep prepare\n```")
        if not any("unknown plan keyword" in error for error in unsupported["errors"]):
            failures.append(f"Unsupported TypeScript plan keyword {keyword} was accepted")
    paused = parse_plan("""```plan
intent draft
surface excel
clarify Which target application should receive the summary?
confidence low
```""")
    if paused["errors"] or not paused["plan"]["needs_clarification"]:
        failures.append("Cross-surface clarification must remain supported")

    print(
        json.dumps(
            {"sample": result, "unsafe": unsafe, "cross": cross, "paused": paused,
             "failures": failures},
            indent=2,
            ensure_ascii=False,
        ),
    )
    if failures:
        raise SystemExit(1)


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        _self_test()
    else:
        print(json.dumps(parse_plan(sys.stdin.read()), indent=2, ensure_ascii=False))
