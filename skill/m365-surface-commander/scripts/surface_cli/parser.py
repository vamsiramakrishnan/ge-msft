"""surface_cli.parser — program-level parsing: wraps the authoritative line parser
(`parse_commands`) and adds the expression/pipeline recognition the checker/normalizer need."""

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from parse_commands import (  # noqa: E402
    _each_body_line_error,
    _has_top_level_pipe,
    _is_each_header,
    _parse_each_header,
    _split_pipeline,
    extract_command_block_meta,
    parse_line,
    verified_frame_error,
)
from .types import _AGG
from .generated_language import TRANSFORM_NAMES

__all__ = [
    "parse_line",
    "extract_command_block_meta",
    "verified_frame_error",
    "_is_expr_line",
    "_infer_pipeline_type",
    "_pipeline_law_errors",
    "_analysis_refs",
    "_is_each_header",
    "_parse_each_header",
    "_each_body_line_error",
]


_COLUMN_STAGES = {"filter", "sort", "derive", "sum", "avg", "min", "max"}


def _stage_columns(name, arg):
    """The column names a stage reads: `filter col=…`, `sort col [desc]`, `sum col`,
    `derive new = a + b` (its operands), `select a,b` (its projection)."""
    arg = arg.strip()
    if not arg:
        return []
    if name == "select":
        return [c.strip() for c in arg.split(",") if c.strip()]
    if name == "filter":
        # C1: a predicate may combine comparisons with and / or / not and parentheses; quoted
        # values are data. Collect the column of every comparison.
        cols = []
        unquoted = re.sub(r'"(?:[^"\\]|\\.)*"', '""', arg)
        for simple in re.split(r"\b(?:and|or)\b", unquoted, flags=re.IGNORECASE):
            simple = re.sub(r"^(?:\s*(?:not\b|\())+", "", simple.strip(), flags=re.IGNORECASE).strip()
            m = re.match(r"^(.+?)\s*(?:>=|<=|!=|=|>|<|\s+contains\s+)", simple, re.IGNORECASE)
            if m:
                cols.append(m.group(1).strip())
        return cols
    if name == "sort":
        return [arg.split()[0]]
    if name == "derive":
        m = re.match(r"^\S+\s*=\s*(\S+)\s*[-+*/]\s*(\S+)$", arg)
        return [t for t in (m.group(1), m.group(2)) if m and not re.match(r"^-?\d+(\.\d+)?$", t)] if m else []
    if name in _AGG:
        return [arg.split()[0]]
    return []


def _pipeline_law_errors(rhs: str):
    """Type and column laws over one pipeline (algebra.md): no stage after an aggregate
    (`sum | head` has no rows to take), and after `select C` a later stage may only name columns
    in C (`filter(p, select(C, T))` is valid iff columns(p) ⊆ C)."""
    parts = [p.strip() for p in _split_pipeline(rhs.strip())]
    errors = []
    selected = None
    aggregated_by = None
    for st in parts[1:]:
        tokens = st.split(None, 1)
        if not tokens:
            continue
        name = tokens[0].lower()
        arg = tokens[1] if len(tokens) > 1 else ""
        if aggregated_by is not None:
            errors.append(f'type law: "{name}" expects a Table, but "{aggregated_by}" produced a Number')
            break
        if name in _AGG:
            aggregated_by = name
        if selected is not None and name in _COLUMN_STAGES:
            for col in _stage_columns(name, arg):
                if col not in selected:
                    errors.append(f'column law: "{name}" uses column "{col}" that "select" did not keep ({", ".join(selected)})')
        if name == "select":
            selected = _stage_columns("select", arg)
    return errors


def _analysis_refs(action):
    """Only typed artifact slots resolve variables; SQL, labels and other text stay literal."""
    kind = action.get("kind")
    values = []
    if kind == "query" and isinstance(action.get("inputs"), list):
        values = list(action["inputs"])
        if isinstance(action.get("requiredColumns"), list):
            values.extend(entry.get("input") for entry in action["requiredColumns"]
                          if isinstance(entry, dict))
    elif kind == "reconcile" and isinstance(action.get("spec"), dict):
        values = [action["spec"].get("left"), action["spec"].get("right")]
    elif kind in ("inspect", "filter", "materialize", "remove"):
        values = [action.get("id")]
    return list(dict.fromkeys(value[1:] for value in values
                             if isinstance(value, str) and re.fullmatch(r"\$[A-Za-z_][A-Za-z0-9_]*", value)))


def _is_expr_line(line: str) -> bool:
    """A `let $x = …` binding or a bare pipeline (a top-level ` | `), mirroring expr-grammar."""
    s = line.strip()
    if re.match(r"^let\s", s, re.IGNORECASE):
        return True
    # Quote- and paren-aware, and independent of surrounding spaces: `comment A1 "x | y"` is a
    # command, `read A1:B3|sum x` is a pipeline — exactly as the TS `findTopLevelPipes`.
    return _has_top_level_pipe(s)


def _infer_pipeline_type(rhs: str):
    """Infer the value type (`Table`/`Number`) of a pipeline RHS from its last transform. Returns
    (vtype, source, transform_names, unknown_transforms)."""
    parts = [p.strip() for p in _split_pipeline(rhs.strip())]
    source = parts[0]
    stages = parts[1:]
    # C1 literal sources type as what they carry; a read / $var / table literal is a Table.
    if re.match(r"^-?\d+(?:\.\d+)?$", source):
        vtype = "Number"
    elif source.startswith('"'):
        vtype = "Text"
    else:
        vtype = "Number" if source.lstrip("$").split()[0] in _AGG else "Table"
    names, unknown = [], []
    for st in stages:
        name = st.split()[0].lower() if st.split() else ""
        names.append(name)
        if name not in TRANSFORM_NAMES:
            unknown.append(name)
        vtype = "Number" if name in _AGG else "Table"
    return vtype, source, names, unknown
