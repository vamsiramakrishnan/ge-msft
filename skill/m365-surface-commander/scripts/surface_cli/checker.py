"""surface_cli.checker — analyze(): parse + capability-scope + type a program into a structured
result (bindings, effects, reads, inferred dependencies). Pure; no rendering, no side effects."""

import json
import re

from .parser import (
    parse_line,
    extract_command_block_meta,
    _is_expr_line,
    _infer_pipeline_type,
    _pipeline_law_errors,
    _analysis_refs,
    _is_each_header,
    _parse_each_header,
    _each_body_line_error,
    verified_frame_error,
)
from .types import _parse_range, _overlap
from .generated_language import LANGUAGE_VERSION, PREFLIGHT
from language_manifest import guard_errors

_READ_PHASE_VERBS = set(PREFLIGHT["readPhaseVerbs"])


def _analysis_guard_errors(action):
    errors = []
    kind = action.get("kind")
    if not isinstance(kind, str):
        return ["analysis kind must be a string"]
    for name, schema in PREFLIGHT["analysisGuards"].get(kind, {}).items():
        if name in action:
            errors.extend(guard_errors(action[name], schema, name))
    # Membership of literal handles is statically decidable; aliases need runtime resolution.
    inputs = action.get("inputs") if isinstance(action.get("inputs"), list) else []
    if action.get("kind") == "query" and not errors:
        for entry in action.get("requiredColumns", []):
            if (not entry["input"].startswith("$")
                    and all(isinstance(value, str) and not value.startswith("$") for value in inputs)
                    and entry["input"] not in inputs):
                errors.append("requiredColumns input must also be declared in query inputs")
    return errors


_DEF_HEADER = re.compile(r"^def\s+([A-Za-z_][A-Za-z0-9_-]*)\s*\(([^)]*)\)\s*:\s*$", re.IGNORECASE)
_STRUCTURAL = {"let", "def", "end", "finish"}


def _split_args(text):
    return [m.group(1) if m.group(1) is not None else m.group(2)
            for m in re.finditer(r'"((?:[^"\\]|\\.)*)"|(\S+)', text)]


def _group_recipes(raw_lines, recipes):
    """Pull `def … end` blocks out of a program. Each body line is checked with its parameters
    replaced by a neutral literal (so `head $n` still parses); the recipe's effect set is recorded
    for the call sites. Returns the remaining top-level lines plus any corrective errors."""
    from .generated_language import TRANSFORM_NAMES  # local to avoid an import cycle at load
    from parse_commands import ALL_VERBS
    errors, remaining = [], []
    current = None
    in_each = False  # C1: an `each … end` block is checked by analyze(), not grouped here
    for raw in raw_lines:
        line = raw.strip()
        if current is None and in_each:
            if line.lower() == "end":
                in_each = False
            remaining.append(raw)
            continue
        if current is None and _is_each_header(line):
            in_each = True
            remaining.append(raw)
            continue
        if current is not None:
            if line.lower() == "end":
                recipes[current["name"]] = current
                current = None
                continue
            if _DEF_HEADER.match(line):
                errors.append(f'nested def in "{current["name"]}" is not allowed')
                continue
            if line and not line.startswith("#"):
                current["body"].append(line)
                substituted = line
                for param in current["params"]:
                    substituted = re.sub(r"\$" + re.escape(param) + r"\b", "1", substituted)
                rec = parse_line(substituted)
                if rec and "error" in rec:
                    errors.append(f'recipe "{current["name"]}": {rec["error"]}')
                elif rec and rec.get("verb") not in (None, "done", "help") and rec.get("verb") not in _READ_PHASE_VERBS:
                    verb = rec["verb"]
                    name = f"/{rec['kind']}" if verb == "invoke" else verb
                    target = rec.get("range") or rec.get("cell") or rec.get("selector")
                    current["effects"].append({"verb": name, "target": target, "range": None,
                                               "external": _authority(name.lstrip("/"), verb == "invoke")["approvalClass"] != "in-document",
                                               **_authority(name.lstrip("/"), verb == "invoke"), "refs": []})
            continue
        m = _DEF_HEADER.match(line)
        if m:
            name = m.group(1).lower()
            params = [p.strip().lstrip("$") for p in re.split(r"[,\s]+", m.group(2).strip()) if p.strip()]
            if name in _STRUCTURAL or name in ALL_VERBS or name in TRANSFORM_NAMES:
                errors.append(f'recipe "{name}" shadows a built-in')
            if len(set(params)) != len(params):
                errors.append(f'recipe "{name}" repeats a parameter')
            current = {"name": name, "params": params, "body": [], "effects": []}
            continue
        if line.lower() == "end":
            errors.append("`end` without a matching `def`")
            continue
        remaining.append(raw)
    if current is not None:
        errors.append(f'def "{current["name"]}" is missing a closing `end`')
    return remaining, errors


def _authority(name, specialized=False):
    info = PREFLIGHT["approvalByKind" if specialized else "approvalByVerb"].get(name)
    return info or {"approvalClass": "in-document", "reversible": True}


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
    recipes = {}
    lines, recipe_errors = _group_recipes(inner.splitlines(), recipes)
    errors.extend(recipe_errors)
    in_each = None  # C1: (row_var, table_var) while inside an `each … end` block

    for raw in lines:
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if finish_requested:
            errors.append("finish when=verified must be the final program entry")
            break

        # C1 `each $row in $table: … end` — the body is a template of effects, expanded per row by
        # the runtime BEFORE dry-run (≤ 50 rows, within the write cap). Here: the table must be a
        # bound (non-artifact) value and every body line an effect naming only `$row.<column>`.
        if in_each is not None and line.lower() == "end":
            in_each = None
            continue
        if _is_each_header(line):
            if in_each is not None:
                errors.append(f"nested each inside each ${in_each[0]} is not allowed")
                continue
            header = _parse_each_header(line)
            if "error" in header:
                errors.append(header["error"])
                continue
            row_var, table_var = header["rowVar"], header["tableVar"]
            if table_var not in bound:
                errors.append(f"each: ${table_var} used before it is bound")
            elif table_var in artifact_bound:
                errors.append(f"each: ${table_var} is an artifact; use analyze to read it")
            in_each = (row_var, table_var)
            continue
        if in_each is not None:
            bad = _each_body_line_error(in_each[0], line)
            if bad:
                errors.append(f'each body line "{line}": {bad}')
                continue

        call = line.split(None, 1)[0].lower()
        if call in recipes:
            # A recipe call expands into its (already checked) body at runtime; the preflight
            # checks arity and counts the body's effects once per call, like the TS expansion.
            given = len(_split_args(line.split(None, 1)[1] if " " in line else ""))
            if given != len(recipes[call]["params"]):
                errors.append(f'recipe "{call}" takes {len(recipes[call]["params"])} argument(s), got {given}')
            for eff in recipes[call]["effects"]:
                if capabilities is not None and eff["verb"] not in capabilities and not eff["verb"].startswith("/"):
                    errors.append(f'"{eff["verb"]}" is not in this turn\'s capabilities')
                effects.append(dict(eff))
            continue

        rec = parse_line(line)
        if rec and "error" in rec:
            errors.append(rec["error"])
            continue
        if rec and rec.get("kind") == "verified-finish":
            finish_requested = True
            # verified_frame_error already enforces the TS whole-response frame contract,
            # including its allowed unclosed block ending exactly at the verified finish.
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
                errors.extend(_pipeline_law_errors(rhs))
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
                errors.extend(_pipeline_law_errors(line))
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
                                "external": _authority("write-cells", True)["approvalClass"] != "in-document",
                                **_authority("write-cells", True), "refs": refs})
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
                 "external": _authority(kind, True)["approvalClass"] != "in-document",
                 **_authority(kind, True), "refs": []}
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
            "external": _authority(verb)["approvalClass"] != "in-document",
            **_authority(verb),
            "refs": re.findall(r"\$(\w+)", rec.get("value", "")) if verb == "spill" else [],
        }
        for ref in eff["refs"]:
            if ref not in bound:
                errors.append(f"spill references unbound ${ref}")
            elif ref in artifact_bound:
                errors.append(f"${ref} is an artifact; use analyze materialize to write it")
        if in_each is not None:
            # One template effect per body line; the runtime multiplies it by the row count.
            eff["loop"] = f"${in_each[0]} in ${in_each[1]}"
        effects.append(eff)
    if in_each is not None:
        errors.append(f"each ${in_each[0]} in ${in_each[1]} is missing a closing `end`")

    # Effect→effect dependencies: a later effect whose range overlaps an earlier effect's range
    # depends on it (the derived-range case — table/chart/cf over a spilled region).
    deps = []
    for i, e in enumerate(effects):
        for j in range(i):
            if e["range"] and effects[j]["range"] and _overlap(e["range"], effects[j]["range"]):
                deps.append((i, j))

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
