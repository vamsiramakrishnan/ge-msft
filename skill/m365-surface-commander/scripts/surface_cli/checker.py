"""surface_cli.checker — analyze(): parse + capability-scope + type a program into a structured
result (bindings, effects, reads, inferred dependencies). Pure; no rendering, no side effects."""

import json
import re

from .parser import (
    parse_line,
    extract_command_block_meta,
    _is_expr_line,
    _infer_pipeline_type,
)
from .types import _parse_range, _overlap, _EXTERNAL, _EXTERNAL_KINDS
from .generated_language import LANGUAGE_VERSION

_READ_PHASE_VERBS = {
    "outline", "read", "search", "list", "inspect", "properties", "comments", "attachments",
    "tables", "slides", "neighbors", "context", "open",
    "workspace", "save", "cat", "grep",
}


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
                if src.split()[0] in _READ_PHASE_VERBS:
                    reads.append(src)
                for ref in re.findall(r"\$(\w+)", rhs):
                    if ref not in bound:
                        errors.append(f"${ref} used before it is bound")
                bound.add(name)
                bindings.append({"name": name, "type": vtype, "source": src})
            else:
                vtype, src, _n, unknown = _infer_pipeline_type(line)
                for u in unknown:
                    errors.append(f'unknown transform "{u}"')
                if src.split()[0] in _READ_PHASE_VERBS:
                    reads.append(src)
            continue

        rec = parse_line(line)
        if rec is None:
            continue
        if "error" in rec:
            errors.append(rec["error"])
            continue

        verb = rec["verb"]
        if verb == "analyze":
            action = json.loads(rec["request"])
            kind = action.get("kind")
            if capabilities is not None and "analyze" not in capabilities:
                errors.append('"analyze" is not in this turn\'s capabilities')
            if kind == "materialize":
                if capabilities is not None and not {"set", "grid", "spill", "write-cells"}.intersection(capabilities):
                    errors.append("analysis materialization requires cell-write capability")
                target = action.get("destination")
                if not isinstance(target, str) or not _parse_range(target):
                    errors.append("analysis materialization requires an explicit A1 destination")
                    target = None
                effects.append({"verb": "analyze:materialize", "target": target,
                                "range": _parse_range(target) if target else None,
                                "external": False, "refs": []})
            elif isinstance(kind, str) and kind in {"capture", "query", "reconcile", "inspect", "filter", "remove"}:
                reads.append(line)
            else:
                errors.append("unsupported analysis action; recovery requires an explicit pane action")
            continue
        if verb in _READ_PHASE_VERBS:
            reads.append(line)
            continue
        if verb in ("done", "help"):
            continue

        # ADR-0008 §two-tier — a `/<kind>` invoke is a specialized effect terminal. It is scoped by
        # its KIND (the command name), not a core verb, and has no composable range.
        if verb == "invoke":
            kind = rec["kind"]
            if capabilities is not None and kind not in capabilities:
                errors.append(f"/{kind} is not in this turn's capabilities")
            effects.append(
                {"verb": f"/{kind}", "target": None, "range": None,
                 "external": kind in _EXTERNAL_KINDS, "refs": []}
            )
            continue

        # An effect. Capability-scope it against this turn's signature, if one was supplied.
        if capabilities is not None and verb not in capabilities:
            errors.append(f'"{verb}" is not in this turn\'s capabilities')
        target = rec.get("range") or rec.get("cell") or rec.get("selector")
        eff = {
            "verb": verb,
            "target": target,
            "range": _parse_range(target) if target else None,
            "external": verb in _EXTERNAL,
            "refs": re.findall(r"\$(\w+)", rec.get("value", "")) if verb == "spill" else [],
        }
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
