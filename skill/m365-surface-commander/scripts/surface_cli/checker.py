"""surface_cli.checker — analyze(): parse + capability-scope + type a program into a structured
result (bindings, effects, reads, inferred dependencies). Pure; no rendering, no side effects."""

import json
import re

from .parser import (
    parse_line,
    extract_command_block_meta,
    _is_expr_line,
    _infer_pipeline_type,
    _analysis_refs,
    verified_frame_error,
)
from .types import _parse_range, _overlap, _EXTERNAL, _EXTERNAL_KINDS
from .generated_language import LANGUAGE_VERSION

_READ_PHASE_VERBS = {
    "outline", "read", "search", "list", "inspect", "properties", "comments", "attachments",
    "tables", "slides", "neighbors", "context", "open",
    "workspace", "save", "cat", "grep",
}


def _analysis_guard_errors(action):
    """Check declared guards without pretending to know live row counts or column schemas."""
    if action.get("kind") == "materialize":
        return (["whenNonEmpty must be a boolean"]
                if "whenNonEmpty" in action and not isinstance(action["whenNonEmpty"], bool) else [])
    if action.get("kind") != "query" or "requiredColumns" not in action:
        return []
    required = action["requiredColumns"]
    if not isinstance(required, list) or len(required) > 16:
        return ["requiredColumns must be an array with at most 16 input guards"]
    errors = []
    inputs = action.get("inputs") if isinstance(action.get("inputs"), list) else []
    for entry in required:
        if not isinstance(entry, dict) or set(entry) - {"input", "indices", "exactDecimal"}:
            errors.append("invalid requiredColumns guard fields")
            continue
        if not isinstance(entry.get("input"), str) or not entry["input"]:
            errors.append("requiredColumns needs an input artifact")
        elif (not entry["input"].startswith("$")
              and all(isinstance(value, str) and not value.startswith("$") for value in inputs)
              and entry["input"] not in inputs):
            # Bound aliases can resolve to the same content-addressed artifact. Only concrete
            # mismatches are decidable here; the runtime checks resolved membership again.
            errors.append("requiredColumns input must also be declared in query inputs")
        indices = entry.get("indices")
        if (not isinstance(indices, list) or not 1 <= len(indices) <= 64
                or any(type(index) is not int or not 0 <= index <= 16383 for index in indices)):
            errors.append("requiredColumns indices must be 1-64 zero-based integers from 0 to 16383")
        if "exactDecimal" in entry and not isinstance(entry["exactDecimal"], bool):
            errors.append("requiredColumns exactDecimal must be a boolean")
    return errors


def analyze(program_text: str, capabilities=None):
    """Parse + scope + type a program. Returns a structured result (no rendering, no side effects)."""
    inner, closed = extract_command_block_meta(program_text)
    # CLI input may be an unfenced program. Once a caller supplies a response frame, enforce the
    # same whole-response boundary as the runtime rather than silently compiling its first block.
    frame_error = verified_frame_error(program_text) if inner is not None else None
    if inner is None:
        inner, closed = program_text, True  # accept a bare program (no fence) for the CLI

    errors, bindings, effects, reads = [], [], [], []
    if frame_error:
        errors.append(frame_error)
        inner = ''
    bound = set()
    artifact_bound = set()
    finish_requested = False

    for raw in inner.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if finish_requested:
            errors.append("finish when=verified must be the final program entry")
            break

        rec = parse_line(line)
        if rec and "error" in rec:
            errors.append(rec["error"])
            continue
        if rec and rec.get("kind") == "verified-finish":
            finish_requested = True
            if not closed:
                errors.append("finish when=verified requires a closed cmd fence")
            continue
        if rec and rec.get("kind") == "analysis-binding":
            action = json.loads(rec["request"])
            errors.extend(_analysis_guard_errors(action))
            if capabilities is not None and "analyze" not in capabilities:
                errors.append('"analyze" is not in this turn\'s capabilities')
            for ref in _analysis_refs(action):
                if ref not in artifact_bound:
                    errors.append(f"analysis references unbound artifact ${ref}")
            name = rec["name"]
            if name in bound:
                errors.append(f"binding ${name} already exists")
            bound.add(name)
            artifact_bound.add(name)
            reads.append(line)
            bindings.append({"name": name, "type": "Artifact", "source": rec["request"]})
            continue

        if _is_expr_line(line):
            m = re.match(r"^let\s+\$(\w+)\s*=\s*(.+)$", line, re.IGNORECASE)
            if m:
                name, rhs = m.group(1), m.group(2)
                if name in artifact_bound:
                    errors.append(f"cannot replace artifact binding ${name} with an expression value")
                vtype, src, _names, unknown = _infer_pipeline_type(rhs)
                for u in unknown:
                    errors.append(f'unknown transform "{u}" in ${name}')
                if src.split()[0] in _READ_PHASE_VERBS:
                    reads.append(src)
                for ref in re.findall(r"\$(\w+)", rhs):
                    if ref not in bound:
                        errors.append(f"${ref} used before it is bound")
                    elif ref in artifact_bound:
                        errors.append(f"${ref} is an artifact; use analyze to read it")
                bound.add(name)
                artifact_bound.discard(name)
                bindings.append({"name": name, "type": vtype, "source": src})
            else:
                vtype, src, _n, unknown = _infer_pipeline_type(line)
                for u in unknown:
                    errors.append(f'unknown transform "{u}"')
                if src.split()[0] in _READ_PHASE_VERBS:
                    reads.append(src)
            continue

        if rec is None:
            continue

        verb = rec["verb"]
        if verb == "analyze":
            action = json.loads(rec["request"])
            errors.extend(_analysis_guard_errors(action))
            kind = action.get("kind")
            if capabilities is not None and "analyze" not in capabilities:
                errors.append('"analyze" is not in this turn\'s capabilities')
            refs = _analysis_refs(action)
            for ref in refs:
                if ref not in artifact_bound:
                    errors.append(f"analysis references unbound artifact ${ref}")
            if kind == "materialize":
                if capabilities is not None and not {"set", "grid", "spill", "write-cells"}.intersection(capabilities):
                    errors.append("analysis materialization requires cell-write capability")
                target = action.get("destination")
                if not isinstance(target, str) or not _parse_range(target):
                    errors.append("analysis materialization requires an explicit A1 destination")
                    target = None
                effects.append({"verb": "analyze:materialize", "target": target,
                                "range": _parse_range(target) if target else None,
                                "external": False, "refs": refs})
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
            elif ref in artifact_bound:
                errors.append(f"${ref} is an artifact; use analyze materialize to write it")
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
        # This is a requested runtime condition, never a proof of host-side verification.
        "requestedCompletion": "verified" if finish_requested else None,
    }
