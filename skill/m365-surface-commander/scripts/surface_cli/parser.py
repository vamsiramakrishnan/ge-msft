"""surface_cli.parser — program-level parsing: wraps the authoritative line parser
(`parse_commands`) and adds the expression/pipeline recognition the checker/normalizer need."""

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from parse_commands import parse_line, extract_command_block_meta, verified_frame_error  # noqa: E402
from .types import _AGG
from .generated_language import TRANSFORM_NAMES

__all__ = [
    "parse_line",
    "extract_command_block_meta",
    "verified_frame_error",
    "_is_expr_line",
    "_infer_pipeline_type",
    "_analysis_refs",
]


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
