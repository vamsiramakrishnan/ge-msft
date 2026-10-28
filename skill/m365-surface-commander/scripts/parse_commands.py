#!/usr/bin/env python3
"""
parse_commands.py — extract the ```cmd fenced block from a reply and parse each command line into a
structured record (verb + args), emitting short corrective errors for malformed lines.

Use this to validate or inspect a command block before relying on it. The Office add-in applies the
authoritative parsing when it runs the commands; this is a lightweight, dependency-free checker.

Usage:
  echo '```cmd
  read Sales!C2:C7
  set Sales!F2 =C2-D2
  ```' | python3 parse_commands.py
  python3 parse_commands.py --self-test
"""

import json
import re
import sys
from pathlib import Path

# Structural argument parsing stays here; vocabulary and advisory semantics are generated.
from language_manifest import load_manifest, ManifestError
_MANIFEST_PATH = Path(__file__).with_name("m365-cli-1.0.json")

# The write verbs `parse_line` has an explicit arg-parsing arm for. The manifest gives the verb SET
# and verb→kind map; the per-verb arg grammar is logic the manifest doesn't encode, so it lives here.
# The self-test asserts this equals the manifest's write verbs — adding a verb to the manifest without
# a Python arm here is a caught drift, not a silent "unhandled verb".
HANDLED_WRITE_VERBS = {
    "set", "grid", "suggest", "comment", "format", "reply",
    "slide", "page", "mail", "post", "compose",
    "shape",
    # ADR-0007 host-native Excel kinds — table/chart/cf take a literal range + props
    # (the `format`-style grammar); spill is the table→grid composition sink (`set` widened).
    "table", "chart", "cf", "spill",
}

def _load_language():
    return load_manifest(_MANIFEST_PATH, "m365-cli/1.0")

try:
    _LANGUAGE = _load_language()
except ManifestError as error:
    raise SystemExit(str(error)) from error

READ_VERBS = set(_LANGUAGE["verbs"]["read"])
WORKSPACE_VERBS = set(_LANGUAGE["verbs"]["workspace"])
CONTROL_VERBS = set(_LANGUAGE["verbs"]["control"])
WRITE_VERBS = set(_LANGUAGE["verbs"]["write"])
TRANSFORM_NAMES = set(_LANGUAGE["transforms"])
EFFECT_VERBS = set(_LANGUAGE["effectVerbs"])
ACTUATION_KINDS = set(_LANGUAGE["actuationKinds"])
COMMAND_HELP = _LANGUAGE["commandHelp"]
CAPABILITY_REGISTRY = _LANGUAGE["capabilityRegistry"]
LANGUAGE_VERSION = _LANGUAGE["version"]
PREFLIGHT = _LANGUAGE["preflight"]
# Default per-turn limits the runtime enforces (absent only in a bundle predating the block).
LIMITS = dict(_LANGUAGE.get("limits") or {})
CONTEXT_KINDS = set(PREFLIGHT["contextKinds"])
CONTEXT_HINTS = set(PREFLIGHT["contextHints"])
ANALYSIS_BINDING_KINDS = frozenset(PREFLIGHT["analysisBindingKinds"])
ALL_VERBS = READ_VERBS | WORKSPACE_VERBS | CONTROL_VERBS | WRITE_VERBS

_FENCE = re.compile(r"```cmd[^\S\n]*\r?\n([\s\S]*?)```", re.IGNORECASE)
# Lenient fallback: an opening ```cmd fence the model never closed — take to end of message.
_FENCE_OPEN = re.compile(r"```cmd[^\S\n]*\r?\n([\s\S]*)$", re.IGNORECASE)
_QUOTED = re.compile(r'"((?:[^"\\]|\\.)*)"')


def extract_command_block(text: str):
    """Return the inner text of the first ```cmd fence, or None (→ re-prompt, not an error).

    Tolerates an unclosed fence: if the model opened ```cmd but never emitted the closing ```,
    we take everything after the opener (a frequent real-world failure mode worth surviving).

    Thin wrapper over {@link extract_command_block_meta} that drops the closed/open flag, kept for
    callers that only need the inner text.
    """
    inner, _closed = extract_command_block_meta(text)
    return inner


def strip_follow_up_block(text: str) -> str:
    """Mirror of the TS `stripFollowUpBlock`: drop the engine's trailing ```suggested_questions
    block, only when it is the last block and directly follows a closed fence."""
    lines = text.rstrip().split("\n")

    def bare(line):
        return line is not None and re.fullmatch(r"```[^\S\n]*", line.rstrip()) is not None

    if len(lines) < 3 or not bare(lines[-1]):
        return text
    opener = len(lines) - 2
    while opener >= 0 and not lines[opener].startswith("```"):
        opener -= 1
    if opener < 0 or not re.fullmatch(r"```suggested_questions[^\S\n]*", lines[opener].rstrip(), re.I):
        return text
    above = opener - 1
    while above >= 0 and lines[above].strip() == "":
        above -= 1
    if above < 0 or not bare(lines[above]):
        return text
    offset = sum(len(line) + 1 for line in lines[:opener])
    return text[:offset].rstrip()


def extract_command_block_meta(text: str):
    """Return `(inner, closed)`: the inner text of the first ```cmd fence (trimmed) and whether the
    fence was properly CLOSED with a trailing ```.

    `(None, False)` ⇒ no ```cmd fence at all (→ re-prompt, not an error). An UNCLOSED fence still
    yields its inner text but with `closed=False`, so the actuation gate can refuse to treat a
    truncated/streaming block as complete (review Finding #8/#10 — a half-emitted block must never
    be honored as `done`).
    """
    m = _FENCE.search(text)
    if m:
        return m.group(1).strip(), True
    m = _FENCE_OPEN.search(text)
    if m:
        # drop a dangling trailing fence if present
        return re.sub(r"```\s*$", "", m.group(1)).strip(), False
    return None, False


def _scan_quoted(rest: str):
    m = _QUOTED.match(rest.strip())
    if not m:
        return None
    val = _unescape(m.group(1))
    return val, rest.strip()[m.end():].strip()


_QUOTED_SHEET = r"'(?:[^']|'')*'(?:[^\s']+|'(?:[^']|'')*')*"


def _split_first_arg(rest: str):
    """Split the first argument while preserving single-quoted sheet names.

    Excel selectors such as `'Daily schedule'!C5:I23` must stay one token even though the
    sheet name contains spaces.
    """
    s = rest.lstrip()
    if not s:
        return None
    m = re.match(r'"([^"]*)"(\S*)|(' + _QUOTED_SHEET + r')|(\S+)', s)
    if not m:
        return None
    if m.group(1) is not None:
        arg = f"{m.group(1)}{m.group(2) or ''}"
    else:
        arg = m.group(3) if m.group(3) is not None else m.group(4)
    return arg, s[m.end():]


def _parse_grid_body(body: str):
    normalized = body.replace("\\n", "\n").replace("\\t", "\t")
    rows = []
    for row in re.split(r"\r?\n", normalized):
        cells = [cell.strip() for cell in row.split("\t")]
        if any(cell != "" for cell in cells):
            rows.append(cells)
    return rows


# Split a verb's argument string into positional tokens + key=value props, keeping `key="quoted
# value"` (with spaces) intact — mirrors the TS `tokenizeArgs` used by table/chart/cf. A
# `key="quoted"` / `key=bare` match yields a prop; anything else is positional (a range, a chart
# type, a bare CF mode like `databar`).
# Double-quoted values honour \", \\ and \n escapes (mirrors tokenizeArgs in command-grammar.ts).
# A token starting with a single-quoted sheet name runs on through later quoted areas
# ('A B'!A1,'A B'!C1); mirrors QUOTED_SHEET_TOKEN in command-grammar.ts.
_TOKENIZE = re.compile(
    r'(\w[\w-]*)="((?:[^"\\]|\\.)*)"|(\w[\w-]*)=(' + _QUOTED_SHEET + r'|\S+)|"((?:[^"\\]|\\.)*)"|('
    + _QUOTED_SHEET
    + r'|\S+)'
)


def _unescape(value: str) -> str:
    # Decode only supported JSON escapes; unknown backslash sequences remain literal.
    return re.sub(r'\\u[dD][89aAbB][0-9a-fA-F]{2}\\u[dD][c-fC-F][0-9a-fA-F]{2}|\\(u[0-9a-fA-F]{4}|["\\ntrbf])', lambda m: json.loads('"' + m.group(0) + '"'), value)


def _tokenize_args(rest: str):
    positional = []
    props = {}
    for m in _TOKENIZE.finditer(rest):
        if m.group(1) is not None:
            props[m.group(1)] = _unescape(m.group(2))
        elif m.group(3) is not None:
            props[m.group(3)] = m.group(4)
        elif m.group(5) is not None:
            positional.append(_unescape(m.group(5)))
        else:
            positional.append(m.group(6))
    return positional, props


def _parse_json_format(rest: str, usage: str):
    """Mirror the TS parser: flatten a model's JSON format facets into key=value props."""
    brace = rest.index("{")
    try:
        obj = json.loads(rest[brace:])
    except ValueError:
        return {"error": "format JSON is malformed (usage: format <range> k=v)"}
    if not isinstance(obj, dict):
        return {"error": usage}
    rng = rest[:brace].strip() or (obj["range"] if isinstance(obj.get("range"), str) else "")
    if not rng or re.search(r"\s", rng):
        return {"error": usage}
    style_value = obj.get("style", obj.get("format"))
    style = style_value if isinstance(style_value, dict) else obj
    props = {}

    def field(v, key):
        return v.get(key) if isinstance(v, dict) else None

    def put(key, v):
        if isinstance(v, bool):
            v = "true" if v else "false"
        elif isinstance(v, (int, float)):
            v = str(v)
        if isinstance(v, str) and key not in props:
            props[key] = v

    def first(*values):
        return next((v for v in values if v is not None), None)

    for key in ("bold", "italic", "fill", "numberFormat", "fontColor", "borderColor"):
        put(key, style.get(key))
    put("bold", field(style.get("font"), "bold"))
    put("italic", field(style.get("font"), "italic"))
    put("fontColor", field(style.get("font"), "color"))
    put("fontColor", style.get("color"))
    put("fill", field(style.get("fill"), "color"))
    put("align", first(style.get("horizontalAlignment"), field(style.get("align"), "horizontal"), style.get("align")))
    put("valign", first(style.get("verticalAlignment"), field(style.get("align"), "vertical")))
    put("border", first(field(style.get("border"), "style"), style.get("border")))
    put("borderColor", field(style.get("border"), "color"))
    if not props:
        return {"error": usage}
    return {"verb": "format", "range": rng, "props": props}


def _is_effect_expr(value: str) -> bool:
    """Mirror the TS `parseEffectArg` accept-set: an effect-arg EXPRESSION is a bare `$var` or a
    parenthesized pipeline `( ... )`. A literal (anything else) is rejected by `spill`.

    Like the TS parser, strip a leading assignment `=` ONLY when followed by whitespace (the
    `spill <range> = (expr)` form) — a `=formula` with no space stays a literal.
    """
    v = value.strip()
    if re.match(r"^=\s", v):
        v = v[1:].strip()
    return v.startswith("$") or v.startswith("(")


def _levenshtein(a: str, b: str) -> int:
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _nearest(word: str, candidates):
    """The closest candidate within the TS grammar's edit-distance threshold
    (≤ min(3, max(1, ⌈len/2⌉))), so both parsers suggest the same correction, or None."""
    best, best_d = None, None
    for cand in sorted(candidates):
        d = _levenshtein(word, cand)
        if best_d is None or d < best_d:
            best, best_d = cand, d
    threshold = min(3, max(1, -(-len(word) // 2)))
    return best if best is not None and best_d <= threshold else None


def _did_you_mean(verb: str):
    near = _nearest(verb, ALL_VERBS)
    return f" — did you mean '{near}'?" if near else ""


def _split_pipeline(value: str):
    parts = []
    buf = []
    quote = False
    escape = False
    depth = 0
    for ch in value:
        if escape:
            buf.append(ch)
            escape = False
            continue
        if ch == "\\" and quote:
            buf.append(ch)
            escape = True
            continue
        if ch == '"':
            quote = not quote
            buf.append(ch)
            continue
        if not quote and ch == "(":
            depth += 1
            buf.append(ch)
            continue
        if not quote and ch == ")":
            depth = max(0, depth - 1)
            buf.append(ch)
            continue
        if not quote and depth == 0 and ch == "|":
            parts.append("".join(buf).strip())
            buf = []
            continue
        buf.append(ch)
    parts.append("".join(buf).strip())
    return parts


def _has_top_level_pipe(value: str) -> bool:
    return len(_split_pipeline(value)) > 1


# C1 literal sources (TS parity: expr-grammar.ts `parseSource`): a number, ONE double-quoted
# string, or `table [[header…], [row…], …]` (JSON rows, header first, cells stringified).
_NUMBER_LITERAL = re.compile(r"^-?\d+(?:\.\d+)?$")
MAX_LITERAL_TABLE_CELLS = 2000


def _parse_table_literal(text: str):
    usage = 'table literal usage: table [["col1","col2"],["a","1"],["b","2"]] (first row = header)'
    try:
        parsed = json.loads(text)
    except ValueError:
        return {"error": f"table literal is not valid JSON — {usage}"}
    if not isinstance(parsed, list) or not parsed or not all(isinstance(r, list) for r in parsed):
        return {"error": f"table literal must be a non-empty array of row arrays — {usage}"}

    def cell(c):
        if c is None:
            return ""
        if isinstance(c, bool):
            return "true" if c else "false"
        if isinstance(c, (str, int, float)):
            return c if isinstance(c, str) else str(c)
        return None

    header = [cell(c) for c in parsed[0]]
    if not header or any(c is None or not c.strip() for c in header):
        return {"error": f"table literal header row needs non-empty column names — {usage}"}
    columns = [c.strip() for c in header]
    if len(set(columns)) != len(columns):
        return {"error": "table literal header has duplicate column names"}
    if len(parsed) * len(columns) > MAX_LITERAL_TABLE_CELLS:
        return {"error": f"table literal too large (max {MAX_LITERAL_TABLE_CELLS} cells) — read it from the document instead"}
    rows = []
    for i, raw in enumerate(parsed[1:], start=2):
        cells = [cell(c) for c in raw]
        if any(c is None for c in cells):
            return {"error": f"table literal row {i} has a non-scalar cell — {usage}"}
        row = cells[: len(columns)] + [""] * max(0, len(columns) - len(cells))
        rows.append(row)
    return {"src": "table", "columns": columns, "rows": rows}


def _parse_pipeline_source(head: str):
    if head == "outline":
        return {"src": "outline"}
    if head.startswith("$"):
        if re.match(r"^\$[A-Za-z_][A-Za-z0-9_]*$", head):
            return {"src": "var", "name": head[1:]}
        return {"error": f"invalid pipeline variable '{head}'"}
    if _NUMBER_LITERAL.match(head):
        return {"src": "number", "value": float(head) if "." in head else int(head)}
    if head.startswith('"'):
        quoted = _scan_quoted(head)
        if not quoted or quoted[1].strip():
            return {"error": 'text literal source must be ONE double-quoted string — e.g. let $s = "hello"'}
        return {"src": "text", "value": quoted[0]}
    parts = head.split(None, 1)
    src = parts[0].lower()
    arg = parts[1].strip() if len(parts) == 2 else ""
    if src == "read":
        # A bare `read` pages the document (TS parity): `read | count` is a valid pipeline.
        return {"src": "read", "selector": arg}
    if src == "table" and arg.startswith("["):
        return _parse_table_literal(arg)
    if len(parts) != 2:
        return {"error": "pipeline needs a source — e.g. read <selector> | <transform>"}
    if src == "search":
        if not arg:
            return {"error": "search pipeline source needs text"}
        return {"src": "search", "text": arg.strip('"')}
    if src in EFFECT_VERBS:
        return {"error": "effects can't be composed yet — emit them as standalone commands (Phase 2)"}
    return {"error": f'unknown pipeline source "{parts[0]}" — use read <selector>, search <text>, outline, $var, a number, a "text" literal, or table [[…]]'}


# C1 `each $row in $table: … end` — the bounded loop. Mirrors packages/contracts/src/each-grammar.ts:
# the body is kept verbatim; each line must start with an effect verb and may reference the row
# only as `$row.<column>`. Expansion (and column checks) happen in the runtime.
_EACH_HEADER_RE = re.compile(r"^each\s+\$([A-Za-z_][A-Za-z0-9_]*)\s+in\s+\$([A-Za-z_][A-Za-z0-9_]*)\s*:$", re.IGNORECASE)
_ROW_REF_RE = re.compile(r"\$([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)")
EACH_HEADER_USAGE = "usage: each $row in $table:  …effect lines…  end"


def _is_each_header(line: str) -> bool:
    return re.match(r"^each(?:\s|$)", line.strip(), re.IGNORECASE) is not None


def _parse_each_header(line: str):
    m = _EACH_HEADER_RE.match(line.strip())
    if not m:
        return {"error": f"malformed each header — {EACH_HEADER_USAGE}"}
    if m.group(1) == m.group(2):
        return {"error": f"each needs distinct names for the row and the table — {EACH_HEADER_USAGE}"}
    return {"rowVar": m.group(1), "tableVar": m.group(2)}


def _each_body_line_error(row_var: str, line: str):
    head = line.split(None, 1)[0].lower()
    if head not in EFFECT_VERBS:
        return f"each body may only contain effect commands ({', '.join(sorted(EFFECT_VERBS))}) — got \"{head}\""
    for m in _ROW_REF_RE.finditer(line):
        if m.group(1) != row_var:
            return f"unknown row variable ${m.group(1)} — the loop variable is ${row_var}"
    if re.search(r"\$" + re.escape(row_var) + r"(?![A-Za-z0-9_.])", line):
        return f"use ${row_var}.<column> to reference a cell — a bare ${row_var} has no value"
    return None


def _collect_each_block(header: str, lines, start: int):
    """Group an each block whose header is `header` and whose body starts at lines[start]. Returns
    (record, next_index) — the record is the block or a corrective error (then the scan skips to
    the next `end` so the rest of the program still parses)."""
    def skip_to_end(i):
        while i < len(lines):
            if lines[i].strip().lower() == "end":
                return i + 1
            i += 1
        return i

    parsed = _parse_each_header(header)
    if "error" in parsed:
        return parsed, skip_to_end(start)
    row_var, table_var = parsed["rowVar"], parsed["tableVar"]
    body = []
    i = start
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        if not line or line.startswith("#"):
            continue
        if line.lower() == "end":
            if not body:
                return {"error": f"each ${row_var} in ${table_var} has an empty body"}, i
            return {"kind": "each-block", "rowVar": row_var, "tableVar": table_var, "body": body}, i
        if _is_each_header(line) or re.match(r"^def\s", line, re.IGNORECASE):
            nested = "each" if _is_each_header(line) else "def"
            return {"error": f"nested {nested} inside each ${row_var} is not allowed"}, skip_to_end(i)
        bad = _each_body_line_error(row_var, line)
        if bad:
            return {"error": f'each body line "{line}": {bad}'}, skip_to_end(i)
        body.append(line)
    return {"error": f"each ${row_var} in ${table_var} is missing a closing `end`"}, i


def _parse_pipeline(value: str):
    segments = _split_pipeline(value)
    if not segments or not segments[0]:
        return {"error": "pipeline needs a source — e.g. read <selector> | <transform>"}
    source = _parse_pipeline_source(segments[0])
    if "error" in source:
        return source
    stages = []
    for seg in segments[1:]:
        if not seg:
            return {"error": "empty pipeline stage (a `|` with nothing after it)"}
        parts = seg.split(None, 1)
        name = parts[0].lower()
        if name not in TRANSFORM_NAMES:
            return {"error": f"unknown transform '{parts[0]}'"}
        stages.append({"name": name, "arg": parts[1].strip() if len(parts) > 1 else ""})
    return {"kind": "pipeline", "source": source, "stages": stages}


def _parse_let(line: str):
    m = re.match(r"^let\s+(\$[A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$", line, re.IGNORECASE)
    if not m:
        return {"error": "let needs a $name and a pipeline — usage: let $name = <source> | <transform> ..."}
    pipeline = _parse_pipeline(m.group(2).strip())
    if "error" in pipeline:
        return pipeline
    return {"kind": "let", "name": m.group(1)[1:], "pipeline": pipeline}


def _validate_workspace_name(name: str):
    if not name:
        return "workspace artifact name cannot be empty"
    if not re.match(r"^[A-Za-z0-9][A-Za-z0-9._/-]*$", name):
        return "workspace artifact name must start with a letter/number and contain only letters, numbers, ., _, -, or /"
    if name.startswith("/") or name.endswith("/") or ".." in name.split("/"):
        return 'workspace artifact name must be relative and cannot contain ".."'
    return None


def _parse_workspace_source(value: str):
    src = value.strip()
    if not src:
        return {"error": "save source is empty"}
    if src == "outline":
        return {"src": "outline"}
    if src.startswith('"'):
        quoted = _scan_quoted(src)
        if not quoted or quoted[1].strip():
            return {"error": 'save literal source must be one quoted string — e.g. save note.md = "text"'}
        return {"src": "literal", "text": quoted[0]}
    if src.startswith("(") and src.endswith(")"):
        pipeline = _parse_pipeline(src[1:-1].strip())
        if "error" in pipeline:
            return pipeline
        return {"src": "expr", "expr": pipeline}
    parts = src.split(None, 1)
    if len(parts) == 2 and parts[0].lower() == "read":
        return {"src": "read", "selector": parts[1].strip()}
    if len(parts) == 2 and parts[0].lower() == "search":
        text = parts[1].strip()
        if not text:
            return {"error": "save search source needs text"}
        return {"src": "search", "text": text.strip('"')}
    return {
        "error": 'save needs a source — usage: save <name> = read <selector> | search <text> | outline | "literal" | ($pipeline)'
    }


def _parse_invoke(verb: str, rest: str):
    """ADR-0008 §two-tier — `/<kind> positional… key=value…` (the specialized surface). The kind must
    be an ActuationKind from the catalogue (loaded from the manifest); an unknown kind yields a
    did-you-mean. Per-surface availability is checked downstream, not here."""
    import difflib

    kind = verb[1:].lower()
    if not kind:
        return {"error": "a / command needs a capability name — usage: /<capability> key=value ..."}
    if kind not in ACTUATION_KINDS:
        near = difflib.get_close_matches(kind, sorted(ACTUATION_KINDS), n=1)
        tail = f" — did you mean '/{near[0]}'?" if near else ""
        return {"error": f"unknown capability '/{kind}'{tail} (the / surface names an ActuationKind)"}
    positional, props = _tokenize_args(rest)
    return {"verb": "invoke", "kind": kind, "props": props, "args": positional}


def _reject_json_constant(_value):
    raise ValueError("Non-JSON numeric constant")


def verified_frame_error(text: str):
    """A completion-bearing model response has one entire, closed frame. Inspect all blocks."""
    if not any(re.match(r"^finish(?:\s|$)", line.strip(), re.IGNORECASE)
               for line in text.splitlines()):
        return None
    # Mirrors the TS grammar: one cmd block and no other fence; prose around it is fine. An
    # unclosed block counts only when its last line is the finish (nothing after it can be lost).
    fences = [line.strip() for line in text.split("\n") if line.strip().startswith("```")]
    embedded = text.count("```") != len(fences)
    last = text.strip().split("\n")[-1].strip() if text.strip() else ""
    opens_cmd = bool(fences) and re.fullmatch(r"```cmd", fences[0], re.IGNORECASE) is not None
    closed = (not embedded and opens_cmd and (
        (len(fences) == 2 and fences[1] == "```")
        or (len(fences) == 1 and re.fullmatch(r"finish\s+when=verified", last, re.IGNORECASE))))
    if closed:
        return None
    if len(fences) <= 1 and not embedded:
        return "finish when=verified requires a closed cmd fence"
    return "finish when=verified requires exactly one cmd fence with no embedded fences"



def _parse_analysis_binding(line: str):
    match = re.match(r"^let\s+([^=]*)=\s*analyze(?:\s+(.*))?$", line, re.IGNORECASE)
    if not match:
        return None
    name = match.group(1).strip()
    if not re.fullmatch(r"\$[A-Za-z_][A-Za-z0-9_]{0,63}", name):
        return {"error": "analysis binding needs a $-prefixed name — usage: let $name = analyze <JSON action>"}
    command = parse_line("analyze " + (match.group(2) or ""))
    if "error" in command:
        return command
    action = json.loads(command["request"])
    kind = action.get("kind")
    if not isinstance(kind, str) or kind not in ANALYSIS_BINDING_KINDS:
        return {"error": "analysis bindings require capture, query, reconcile, filter, or inspect; effects and recovery cannot bind artifacts"}
    return {"kind": "analysis-binding", "name": name[1:], "request": command["request"]}


def parse_line(line: str):
    """Parse one command line → a dict record, or {'error': '...'} (the corrective contract)."""
    line = line.strip()
    if not line or line.startswith("#"):
        return None
    parts = line.split(None, 1)
    raw_verb = parts[0]
    verb = raw_verb.lower()
    rest = parts[1].strip() if len(parts) > 1 else ""

    if verb == "finish":
        if re.fullmatch(r"finish\s+when=verified", line, re.IGNORECASE):
            return {"kind": "verified-finish"}
        return {"error": "finish requires exactly when=verified — usage: finish when=verified"}
    if verb == "let":
        binding = _parse_analysis_binding(line)
        if binding is not None:
            return binding
        return _parse_let(line)
    if verb == "each":
        # Only meaningful as a block (grouped by parse_block); a lone header is corrective.
        parsed = _parse_each_header(line)
        return parsed if "error" in parsed else {"error": f"each ${parsed['rowVar']} in ${parsed['tableVar']} is missing a closing `end`"}
    if _has_top_level_pipe(line):
        return _parse_pipeline(line)

    if rest in ("-h", "--help"):
        return {"verb": "help", "topic": raw_verb}

    # ADR-0008 §two-tier — a `/<kind>` line is the SPECIALIZED surface (the long-tail catalogue),
    # dispatched before the core-verb check. The command name IS the ActuationKind (drift-free).
    if raw_verb.startswith("/"):
        return _parse_invoke(verb, rest)

    if verb not in ALL_VERBS:
        return {"error": f"unknown verb '{verb}'{_did_you_mean(verb)} (run 'help')"}

    # No-argument verbs. Consume the FULL line: a trailing token is malformed input, not something
    # to silently drop (review Finding #8/#10 — a dropped `done`-block tail hid real parse failures).
    if verb == "outline":
        if rest:
            return {"error": f"{verb} takes no arguments — got {rest!r} (usage: {verb})"}
        return {"verb": verb}

    if verb == "done":
        if not rest:
            return {"verb": "done"}
        # Never an error: trailing text after done still finishes; only a quoted answer is kept.
        q = _scan_quoted(rest) if rest.strip().startswith('"') else None
        return {"verb": "done", "message": q[0]} if q and q[0].strip() else {"verb": "done"}
    if verb == "help":
        return {"verb": "help", **({"topic": rest} if rest else {})}
    if verb == "read":
        # A wrapping pair of double quotes is not part of an Excel address (TS: command-grammar.ts).
        if len(rest) >= 2 and rest.startswith('"') and rest.endswith('"'):
            rest = rest[1:-1].strip()
        return {"verb": "read", "selector": rest}  # empty ⇒ first bounded document page; follow next to EOF
    if verb == "search":
        if not rest:
            return {"error": "search needs text — usage: search <text>"}
        return {"verb": "search", "text": rest.strip('"')}
    if verb == "ls":
        if not rest:
            return {"error": "ls needs a path — usage: ls <path>, e.g. ls /doc"}
        return {"verb": "ls", "path": rest}
    if verb == "find":
        if not rest:
            return {"error": "find needs a path — usage: find <path> [glob]"}
        tokens = rest.split()
        out = {"verb": "find", "path": tokens[0]}
        if len(tokens) > 1:
            out["glob"] = tokens[1]
        return out
    if verb == "tail":
        if not rest:
            return {"error": "tail needs a path — usage: tail <path> [n]"}
        tokens = rest.split()
        out = {"verb": "tail", "path": tokens[0]}
        if len(tokens) > 1:
            try:
                out["n"] = int(tokens[1])
            except ValueError:
                return {"error": "tail: n must be a number"}
        return out
    if verb == "list":
        if not rest:
            return {"verb": "list"}
        kind = rest.lower()
        # Accept the plural the model often writes (`list comments` -> `comment`).
        if kind not in CONTEXT_KINDS and kind.endswith("s"):
            kind = kind[:-1]
        if kind not in CONTEXT_KINDS:
            supported = ", ".join(sorted(CONTEXT_KINDS))
            return {"error": f"unknown context kind '{rest}' — supported: {supported}"}
        return {"verb": "list", "kind": kind}
    if verb in ("inspect", "properties", "open"):
        if not rest:
            return {"error": f"{verb} needs a ref id or selector — usage: {verb} <refId|selector>"}
        return {"verb": verb, "selector": rest.strip('"')}
    if verb in ("comments", "attachments", "tables", "slides", "neighbors"):
        out = {"verb": verb}
        if rest:
            out["selector"] = rest.strip('"')
        return out
    if verb == "context":
        hints = []
        for raw in rest.split():
            hint = raw.lower()
            if hint not in CONTEXT_HINTS:
                supported = ", ".join(sorted(CONTEXT_HINTS))
                return {"error": f"unknown context hint '{raw}' — supported: {supported}"}
            hints.append(hint)
        return {"verb": "context", "hints": hints}

    if verb == "analyze":
        if not rest or len(rest) > 32768:
            return {"error": "analyze requires a bounded JSON action object"}
        try:
            if not isinstance(json.loads(rest, parse_constant=_reject_json_constant), dict):
                raise ValueError()
        except (ValueError, TypeError):
            return {"error": "analyze requires a JSON action object"}
        return {"verb": "analyze", "request": rest}
    if verb == "workspace":
        return {"verb": "workspace", **({"ref": rest.strip('"')} if rest else {})}

    # `share` has the exact same source grammar as `save` — it only differs in where the runtime
    # persists the result (the cross-surface `/shared` Graph app-folder store, not local `/work`).
    if verb in ("save", "share"):
        m = re.match(r"^(\S+)\s*=\s*(.+)$", rest)
        if not m:
            return {
                "error": f'{verb} needs a name and source — usage: {verb} <name> = read <selector> | search <text> | outline | "literal" | ($pipeline)'
            }
        name = m.group(1)
        name_error = _validate_workspace_name(name)
        if name_error:
            return {"error": name_error}
        source = _parse_workspace_source(m.group(2))
        if "error" in source:
            return source
        return {"verb": verb, "name": name, "source": source}

    if verb == "cat":
        split = _split_first_arg(rest)
        if not split:
            return {"error": "cat needs an artifact ref or absolute path — usage: cat <name|ws:id|/absolute/path> [head=N]"}
        ref, tail = split[0], split[1].strip()
        out = {"verb": "cat", "ref": ref}
        if tail:
            m = re.fullmatch(r"head=(\d+)", tail)
            if not m:
                return {"error": "cat only accepts optional head=N — usage: cat <name|ws:id> [head=N]"}
            out["head"] = int(m.group(1))
        return out

    if verb == "grep":
        split = _split_first_arg(rest)
        if not split:
            return {"error": 'grep needs an artifact ref or absolute path and pattern — usage: grep <name|ws:id|/absolute/path> "pattern" [context=N]'}
        ref, tail = split[0], split[1].strip()
        if not tail:
            return {"error": 'grep needs an artifact ref or absolute path and pattern — usage: grep <name|ws:id|/absolute/path> "pattern" [context=N]'}
        context = None
        ctx = re.search(r"(?:^|\s)context=(\d+)\s*$", tail)
        if ctx:
            context = int(ctx.group(1))
            tail = tail[: ctx.start()].strip()
        quoted = _scan_quoted(tail)
        if quoted:
            pattern, after = quoted
            if after.strip():
                return {"error": 'grep only accepts a quoted pattern plus optional context=N'}
        else:
            pattern = tail
        if not pattern:
            return {"error": "grep pattern cannot be empty"}
        out = {"verb": "grep", "ref": ref, "pattern": pattern}
        if context is not None:
            out["context"] = context
        return out

    # `cp <src> <dst>` / `mv <src> <dst>` — duplicate/rename a workspace artifact (`/work` only,
    # never touches Office content). `dst` is validated with the same rules as `save`'s name (it
    # names a fresh alias); `src` is a bare ref, resolved by WorkspaceStore.get() at execution.
    if verb in ("cp", "mv"):
        usage = f"{verb} needs a source and destination — usage: {verb} <src> <dst>"
        positional, props = _tokenize_args(rest)
        if len(positional) != 2 or props:
            return {"error": usage}
        src, dst = positional
        name_error = _validate_workspace_name(dst)
        if name_error:
            return {"error": name_error}
        return {"verb": verb, "src": src, "dst": dst}

    # `rm <name|ws:id>` — delete a workspace artifact (`/work` only, never touches Office content).
    if verb == "rm":
        usage = "rm needs an artifact ref — usage: rm <name|ws:id>"
        positional, props = _tokenize_args(rest)
        if len(positional) != 1 or props:
            return {"error": usage}
        return {"verb": "rm", "name": positional[0]}

    if verb == "set":
        sp = re.search(r"\s", rest)
        if not sp:
            return {"error": "set needs a cell and a value — usage: set <A1|A1:A50> <value|=formula>"}
        cell, value = rest[: sp.start()], rest[sp.start():].strip()
        if not cell or not value:
            return {"error": "set needs a cell and a value — usage: set <A1|A1:A50> <value|=formula>"}
        return {"verb": "set", "cell": cell, "value": value}

    if verb == "grid":
        usage = 'grid needs a range and quoted TSV — usage: grid <range> = "a\\tb\\nc\\td"'
        split = _split_first_arg(rest)
        if not split:
            return {"error": usage}
        rng, tail = split[0], split[1].strip()
        if tail.startswith("="):
            tail = tail[1:].strip()
        quoted = _scan_quoted(tail)
        if not rng or not quoted or quoted[1].strip():
            return {"error": usage}
        cells = _parse_grid_body(quoted[0])
        if not cells or all(all(cell == "" for cell in row) for row in cells):
            return {"error": "grid body is empty — provide at least one non-empty cell"}
        width = len(cells[0])
        if width == 0 or any(len(row) != width for row in cells):
            return {"error": "grid rows must be rectangular — every row needs the same number of cells"}
        return {"verb": "grid", "range": rng, "cells": cells}

    if verb == "suggest":
        first = _scan_quoted(rest)
        if not first:
            return {"error": 'suggest needs two quoted strings — usage: suggest "old" => "new"'}
        old, tail = first
        sep = re.match(r"\s*(=>|->)\s*", tail)
        if not sep:
            return {"error": 'suggest needs two quoted strings — usage: suggest "old" => "new"'}
        tail = tail[sep.end():]
        second = _scan_quoted(tail)
        if not second:
            return {"error": 'suggest needs two quoted strings — usage: suggest "old" => "new"'}
        new, after = second
        # Mirror the TS parser: anything after the closing quote (besides whitespace) is malformed.
        if after.strip():
            return {"error": 'suggest needs two quoted strings — usage: suggest "old" => "new"'}
        return {"verb": "suggest", "oldText": old, "newText": new}

    if verb == "comment":
        # Two accepted forms:
        #   Excel : comment <cell> "note"        (unquoted selector + one quoted note)
        #   Word  : comment "anchor" "note"      (two quoted strings — anchor + note)
        quoted = _QUOTED.findall(rest)
        if rest.lstrip().startswith('"') and len(quoted) >= 2:
            return {"verb": "comment", "selector": quoted[0], "text": quoted[1]}
        q = _QUOTED.search(rest)
        if not q:
            return {"error": 'comment needs a selector and a quoted note — '
                             'usage: comment <cell> "note"  OR  comment "anchor" "note"'}
        selector = rest[: q.start()].strip().strip('"')
        if not selector:
            return {"error": 'comment is missing its note — usage: comment "anchor" "note"'}
        return {"verb": "comment", "selector": selector, "text": q.group(1)}

    if verb == "format":
        usage = "format needs a range and at least one key=value — usage: format <range> k=v ..."
        if "{" in rest:
            return _parse_json_format(rest, usage)
        tokens = [t for t in re.split(r"\s+", rest) if t]
        if not tokens:
            return {"error": usage}
        rng, pairs = tokens[0], tokens[1:]
        if not pairs:
            return {"error": usage}
        # Mirror the TS parser: every trailing token MUST be key=value (split on the first `=`).
        # A bare token is malformed and reported, never silently dropped.
        props = {}
        for pair in pairs:
            eq = pair.find("=")
            if eq <= 0:
                return {"error": f'format expects key=value pairs — got "{pair}" (usage: format <range> k=v)'}
            props[pair[:eq]] = pair[eq + 1:]
        return {"verb": "format", "range": rng, "props": props}

    if verb == "reply":
        m = re.match(r"(\S+)\s+", rest)
        q = _QUOTED.search(rest)
        if not m or not q:
            return {"error": 'reply needs a comment id and a quoted reply — usage: reply <id> "text" [resolve]'}
        out = {"verb": "reply", "commentId": m.group(1), "text": q.group(1)}
        # `resolve` after the quoted text also marks the thread resolved (mirrors parseReply).
        if rest[q.end():].strip() == "resolve":
            out["resolve"] = True
        return out

    if verb in ("page", "compose"):
        qs = _QUOTED.findall(rest)
        if len(qs) < 2:
            return {"error": f'{verb} needs a quoted title/subject and a body'}
        key = "title" if verb == "page" else "subject"
        return {"verb": verb, key: qs[0], "body": qs[1]}

    if verb in ("mail", "post"):
        q = _scan_quoted(rest)
        if not q:
            return {"error": f'{verb} needs a quoted body/text'}
        slot = "body" if verb == "mail" else "text"
        out = {"verb": verb, slot: q[0]}
        # mail "reply" attach="name.docx" text="file text": a file opened with the reply.
        # Mirrors parseMail: a tail is only attach= and text=, both present, after a non-empty body.
        if verb == "mail" and q[1].strip():
            positional, props = _tokenize_args(q[1])
            if positional or set(props) - {"attach", "text"} or "attach" not in props or "text" not in props:
                return {"error": 'mail needs a quoted body, then optionally attach="name.docx" text="file text"'}
            if q[0] == "":
                return {"error": "mail body cannot be empty"}
            if not props["attach"].strip():
                return {"error": "mail attach= needs a file name"}
            out["attachment"] = {"name": props["attach"], "text": props["text"]}
        return out

    if verb == "slide":
        at = None
        m = re.match(r"^at=(\S*)\s+", rest.strip(), re.I) or re.search(r"\s+at=(\S*)$", rest.strip(), re.I)
        if m:
            if not re.fullmatch(r"[1-9]\d*", m.group(1)):
                return {"error": f'slide at= must be a slide number (1 = first), got "{m.group(1)}"'}
            at = int(m.group(1))
            t = rest.strip()
            rest = t[m.end():] if m.start() == 0 else t[: m.start()]
        qs = _QUOTED.findall(rest)
        if not qs:
            return {"error": 'slide needs a quoted title — usage: slide "Title" "bullet" ... [at=N]'}
        out = {"verb": "slide", "title": qs[0], "bullets": qs[1:]}
        if at is not None:
            out["at"] = at
        return out

    if verb == "shape":
        q = _QUOTED.search(rest)
        if not q:
            return {"error": 'shape needs a shape ref and quoted text — usage: shape <pp:shape:slideId:shapeId> "text"'}
        selector = rest[: q.start()].strip()
        if not selector:
            return {"error": 'shape needs a shape ref and quoted text — usage: shape <pp:shape:slideId:shapeId> "text"'}
        after = rest[q.end():].strip()
        if after:
            return {"error": 'shape needs a shape ref and quoted text — usage: shape <pp:shape:slideId:shapeId> "text"'}
        return {"verb": "shape", "selector": selector, "text": q.group(1)}

    # ADR-0007 `table <range> [headers] [name=...]` — promote a range to a native Table.
    if verb == "table":
        positional, props = _tokenize_args(rest)
        if not positional:
            return {"error": "table needs a range — usage: table <range> [headers] [name=...]"}
        rng = positional[0]
        # A bare `headers` flag is sugar for headers=true (the common case).
        if "headers" in positional[1:]:
            props["headers"] = "true"
        return {"verb": "table", "range": rng, "props": props}

    # ADR-0007 `chart <type> <range> [title="…"] [series=rows|columns]` — a chart over a range.
    if verb == "chart":
        usage = ("chart needs a type and a range — usage: chart "
                 "<column|bar|line|pie|scatter|area> <range> [title=\"…\"] [series=rows|columns] "
                 "[x=<col> by=<col> value=<col>]")
        positional, props = _tokenize_args(rest)
        if len(positional) < 2:
            return {"error": usage}
        # The source may put a space after a comma (`C1:C9, G1:G9`); rejoin the areas split there.
        rng = positional[1]
        i = 2
        while rng.endswith(",") and i < len(positional):
            rng += positional[i]
            i += 1
        return {"verb": "chart", "chartType": positional[0].lower(),
                "range": rng, "props": props}

    # ADR-0007 `cf <range> <rule>` — one conditional-format rule. Tolerant of an inline operator
    # (`>1000`), a bare mode (`databar`/`colorscale`), or `top=N`. Only collects props.
    if verb == "cf":
        positional, props = _tokenize_args(rest)
        if not positional:
            return {"error": "cf needs a range and a rule — usage: cf <range> >VALUE [fill=#hex] | "
                             "cf <range> databar|colorscale | cf <range> top=N"}
        rng = positional[0]
        for tok in positional[1:]:
            op = re.match(r"(>=|<=|!=|>|<|=)(.+)$", tok)
            if op:
                props["op"] = op.group(1)
                props["value"] = op.group(2)
            else:
                props[tok.lower()] = "true"  # a bare mode: databar / colorscale
        if not props:
            return {"error": "cf needs a rule — usage: cf <range> >VALUE [fill=#hex] | "
                             "cf <range> databar|colorscale | cf <range> top=N"}
        return {"verb": "cf", "range": rng, "props": props}

    # ADR-0007 §3 `spill <range> = <expr>` — write a composed TABLE as a grid. The range is the
    # first token; the remainder MUST be a `$var` / `( pipeline )` expression — a literal is
    # rejected (spill is the composition sink, not a verbatim writer; use `set` for one cell).
    if verb == "spill":
        usage = "spill needs a range and a table expression — usage: spill <range> = ($rows)"
        sp = re.search(r"\s", rest)
        if not sp:
            return {"error": usage}
        rng, value = rest[: sp.start()], rest[sp.start():].strip()
        if not rng or not value:
            return {"error": usage}
        if not _is_effect_expr(value):
            return {"error": "spill needs a composed table, not a literal — e.g. "
                             "spill Report!A1 = ($rows) (use set for one cell)"}
        return {"verb": "spill", "range": rng, "value": value}

    return {"error": f"unhandled verb '{verb}'"}


def parse_block(model_text: str):
    inner, closed = extract_command_block_meta(model_text)
    if inner is None:
        return {
            "block": None,
            "closed": False,
            "commands": [],
            "note": "no ```cmd fence (re-prompt, not an error)",
        }
    out = []
    finished = False
    lines = inner.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        if not line or line.startswith("#"):
            continue
        if finished:
            out.append({"error": "finish when=verified must be the final program entry"})
            break
        if line.lower() == "end":
            out.append({"error": "`end` without a matching `def` or `each`"})
            continue
        if _is_each_header(line):
            rec, i = _collect_each_block(line, lines, i)
            out.append(rec)
            continue
        rec = parse_line(line)
        if rec is not None:
            out.append(rec)
            finished = rec.get("kind") == "verified-finish"
    requests_finish = any(re.match(r"^finish(?:\s|$)", line.strip(), re.IGNORECASE)
                          for line in model_text.splitlines())
    frame_error = verified_frame_error(model_text)
    if frame_error:
        out.append({"error": frame_error})
    if requests_finish and any("error" in rec for rec in out):
        out = [rec for rec in out if "error" in rec]
    return {"block": inner, "closed": closed, "commands": out}


def block_is_complete(parsed) -> bool:
    """Decide whether a parsed block may be honored as DONE — the actuation gate's fail-closed
    primitive (review Finding #8/#10; README §"What we learned": the add-in must "not honor `done`
    if the same block had parse errors").

    Complete iff ALL of:
      • a ```cmd fence was present and properly CLOSED (a truncated/streaming block is never done);
      • NO line produced a parse error (a single malformed line poisons the whole block); and
      • at least one `done` command is present.

    A block missing `done`, a block with any error, and an unclosed fence each return False — so a
    partially-broken or in-flight block can never trip the gate.
    """
    if parsed.get("block") is None or not parsed.get("closed"):
        return False
    cmds = parsed.get("commands", [])
    if any("error" in c for c in cmds):
        return False
    return any(c.get("verb") == "done" for c in cmds)


def _self_test() -> int:
    # This sample carries BOTH a parse error (`writ-cells`) AND `done` — the fail-closed case from
    # review Finding #8/#10. The block must NOT be reported complete.
    sample = """**thought** I'll read then write.
```cmd
read Sales!C2:C7
set Sales!F2 =C2-D2
suggest "old wording" => "new wording"
comment Sales!A16 "anomalous spike"
format Sales!A16:C16 bold=true fill=#FFF2CC
writ-cells A1 5
done
```"""
    parsed = parse_block(sample)
    print(json.dumps(parsed, indent=2))

    failures = []
    if block_is_complete(parsed):
        failures.append("block with `done` + a parse error was reported complete (fail-open!)")

    # An unclosed fence is never complete even with a clean `done`.
    if block_is_complete(parse_block("```cmd\nread Sales!C2:C7\ndone")):
        failures.append("unclosed fence reported complete")

    # The engine's trailing follow-up block is dropped only after a closed fence.
    closed_cmd = "```cmd\nread Sales!C2:C7\ndone\n```"
    follow = "```suggested_questions\n[\"Add a chart?\"]\n```"
    if strip_follow_up_block(closed_cmd + "\n\n" + follow) != closed_cmd:
        failures.append("trailing suggested_questions block after a closed fence was not dropped")
    unclosed = "```cmd\nread Sales!C2:C7\ndone\n" + follow
    if strip_follow_up_block(unclosed) != unclosed:
        failures.append("follow-up block after an unclosed fence was dropped")

    # A clean, closed, done-bearing block IS complete.
    if not block_is_complete(parse_block("```cmd\nread Sales!C2:C7\ndone\n```")):
        failures.append("clean closed `done` block was NOT reported complete")

    # A no-arg verb with a trailing token is a reported error, not a silent drop.
    if parse_line("done now please") != {"verb": "done"}:
        failures.append("`done now please` did not finish (trailing text is ignored, as in the TS grammar)")

    ctx = parse_line("context analytical full-scope upload-preferred code-execution-preferred")
    if ctx != {
        "verb": "context",
        "hints": ["analytical", "full-scope", "upload-preferred", "code-execution-preferred"],
    }:
        failures.append(f"context hints did not parse: {ctx}")
    if "error" not in (parse_line("context run-python-now") or {}):
        failures.append("unknown context hint did not error")
    if parse_line("list range") != {"verb": "list", "kind": "range"}:
        failures.append("list range did not parse")
    if parse_line("ls /doc") != {"verb": "ls", "path": "/doc"}:
        failures.append("ls /doc did not parse")
    if "error" not in (parse_line("ls") or {}):
        failures.append("bare ls did not error")
    if parse_line("find /work") != {"verb": "find", "path": "/work"}:
        failures.append("find /work did not parse")
    if parse_line("find /work *.tsv") != {"verb": "find", "path": "/work", "glob": "*.tsv"}:
        failures.append("find /work *.tsv did not parse")
    if "error" not in (parse_line("find") or {}):
        failures.append("bare find did not error")
    if parse_line("tail /work/notes.md") != {"verb": "tail", "path": "/work/notes.md"}:
        failures.append("tail /work/notes.md did not parse")
    if parse_line("tail /work/notes.md 20") != {
        "verb": "tail",
        "path": "/work/notes.md",
        "n": 20,
    }:
        failures.append("tail /work/notes.md 20 did not parse")
    if "error" not in (parse_line("tail") or {}):
        failures.append("bare tail did not error")
    if "error" not in (parse_line("tail /work/notes.md abc") or {}):
        failures.append("tail with a non-numeric n did not error")
    if parse_line('read "Branch Performance!A1:F21"') != {"verb": "read", "selector": "Branch Performance!A1:F21"}:
        failures.append("read with a double-quoted selector kept its quotes")
    if parse_line('open "Sales!A1:C9"') != {"verb": "open", "selector": "Sales!A1:C9"}:
        failures.append("open selector did not parse")
    if parse_line('mail "See attached" attach="a.docx" text="Hi"') != {
        "verb": "mail", "body": "See attached", "attachment": {"name": "a.docx", "text": "Hi"}
    }:
        failures.append("mail with attach= did not parse")
    for bad in ('mail "x" attach="a.docx" text="y" extra', 'mail "x" attach="a.docx" text="y" cc="z"',
                'mail "" attach="a.docx" text="y"'):
        if "error" not in (parse_line(bad) or {}):
            failures.append(f"{bad} did not error")
    if "error" not in (parse_line('mail "See attached" attach="a.docx"') or {}):
        failures.append("mail attach= without text= did not error")
    if parse_line("list comments") != {"verb": "list", "kind": "comment"}:
        failures.append("plural list kind did not parse")
    if "error" not in (parse_line("list monster") or {}):
        failures.append("unknown list kind did not error")
    if parse_line("help shape") != {"verb": "help", "topic": "shape"}:
        failures.append("help shape did not parse")
    if parse_line("shape -h") != {"verb": "help", "topic": "shape"}:
        failures.append("shape -h did not parse as targeted help")
    if parse_line("/insert-image -h") != {"verb": "help", "topic": "/insert-image"}:
        failures.append("/insert-image -h did not parse as targeted help")
    shape = parse_line('shape pp:shape:s2:s2-shape-1 "Updated outlook"')
    if shape != {"verb": "shape", "selector": "pp:shape:s2:s2-shape-1", "text": "Updated outlook"}:
        failures.append(f"shape did not parse: {shape}")
    if "error" not in (parse_line("shape pp:shape:s2:s2-shape-1") or {}):
        failures.append("malformed shape did not error")
    bound = parse_line("let $top = read Sales!A1:D20 | filter Quarter=Q3 | head 5")
    if bound.get("kind") != "let" or bound.get("name") != "top":
        failures.append(f"let pipeline did not parse: {bound}")
    pipe = parse_line("read Sales!A1:D20 | count")
    if pipe.get("kind") != "pipeline" or len(pipe.get("stages", [])) != 1:
        failures.append(f"bare pipeline did not parse: {pipe}")
    if "error" not in (parse_line("let top = read Sales!A1:D20 | count") or {}):
        failures.append("malformed let did not error")
    if "error" not in (parse_line("read Sales!A1:D20 | explode") or {}):
        failures.append("unknown transform did not error")

    grid = parse_line("grid 'Daily schedule'!C5:D6 = \"Monday\\tTuesday\\nDeep Work\\tMusic Lesson\"")
    if grid != {
        "verb": "grid",
        "range": "'Daily schedule'!C5:D6",
        "cells": [["Monday", "Tuesday"], ["Deep Work", "Music Lesson"]],
    }:
        failures.append(f"grid did not parse: {grid}")
    if "error" not in (parse_line('grid Report!A1:B2 = "A\\tB\\nC"') or {}):
        failures.append("ragged grid did not error")

    # ADR-0008 §4 drift gate: every manifest write verb MUST have a parse arm here (and vice-versa),
    # so growing the manifest without a Python arm is caught — never a silent "unhandled verb".
    if HANDLED_WRITE_VERBS != WRITE_VERBS:
        missing = sorted(WRITE_VERBS - HANDLED_WRITE_VERBS)
        extra = sorted(HANDLED_WRITE_VERBS - WRITE_VERBS)
        failures.append(f"manifest/parse-arm drift — manifest-only: {missing}; arm-only: {extra}")

    # ADR-0008 §two-tier: the /<kind> specialized surface parses a catalogue kind into an invoke,
    # and rejects an unknown kind with a did-you-mean.
    inv = parse_line('/insert-image base64=AAA alt="chart"')
    if inv != {"verb": "invoke", "kind": "insert-image", "props": {"base64": "AAA", "alt": "chart"}, "args": []}:
        failures.append(f"/insert-image did not parse to an invoke: {inv}")
    if ACTUATION_KINDS and "error" not in (parse_line("/insert-imag base64=AAA") or {}):
        failures.append("unknown /kind did not error")
    if "shape" in WRITE_VERBS and "shape" not in COMMAND_HELP:
        failures.append("generated commandHelp is missing shape")
    if parse_line("save schedule.tsv = read 'Daily schedule'!B3:I53") != {
        "verb": "save",
        "name": "schedule.tsv",
        "source": {"src": "read", "selector": "'Daily schedule'!B3:I53"},
    }:
        failures.append(f"save read did not parse: {parse_line('save schedule.tsv = read ' + chr(39) + 'Daily schedule' + chr(39) + '!B3:I53')}")
    if parse_line('share note.txt = "hello"') != {
        "verb": "share",
        "name": "note.txt",
        "source": {"src": "literal", "text": "hello"},
    }:
        failures.append(f"share literal did not parse: {parse_line('share note.txt = ' + chr(34) + 'hello' + chr(34))}")
    if parse_line('cat schedule.tsv head=12') != {"verb": "cat", "ref": "schedule.tsv", "head": 12}:
        failures.append("cat head did not parse")
    if parse_line('grep schedule.tsv "Deep Work" context=1') != {
        "verb": "grep",
        "ref": "schedule.tsv",
        "pattern": "Deep Work",
        "context": 1,
    }:
        failures.append("grep context did not parse")
    if "error" not in (parse_line("save ../bad = outline") or {}):
        failures.append("unsafe workspace artifact name did not error")
    if parse_line("cp a.tsv b.tsv") != {"verb": "cp", "src": "a.tsv", "dst": "b.tsv"}:
        failures.append("cp a.tsv b.tsv did not parse")
    if parse_line("mv a.tsv b.tsv") != {"verb": "mv", "src": "a.tsv", "dst": "b.tsv"}:
        failures.append("mv a.tsv b.tsv did not parse")
    if parse_line("rm a.tsv") != {"verb": "rm", "name": "a.tsv"}:
        failures.append("rm a.tsv did not parse")
    if "error" not in (parse_line("cp a.tsv") or {}):
        failures.append("cp with a missing destination did not error")
    if "error" not in (parse_line("mv a.tsv") or {}):
        failures.append("mv with a missing destination did not error")
    if "error" not in (parse_line("rm") or {}):
        failures.append("bare rm did not error")
    if "error" not in (parse_line("cp a.tsv ../bad") or {}):
        failures.append("unsafe cp destination name did not error")
    if "error" not in (parse_line("mv a.tsv ../bad") or {}):
        failures.append("unsafe mv destination name did not error")

    if failures:
        print("SELF-TEST FAIL", file=sys.stderr)
        for f in failures:
            print(f"  - {f}", file=sys.stderr)
        return 1
    assert parse_line('analyze {"kind":"capture","range":"S!A1:C2"}')["verb"] == "analyze"
    assert "error" in parse_line('analyze {"kind":"query","value":NaN}')
    assert "error" in parse_line('analyze []')
    assert parse_line('let $x = analyze {"kind":"capture","range":"S!A1:C2"}') == {
        "kind": "analysis-binding", "name": "x", "request": '{"kind":"capture","range":"S!A1:C2"}'
    }
    for kind in ("materialize", "remove", "recovery", "undo", "resume", "forget", None, []):
        assert "error" in parse_line("let $x = analyze " + json.dumps({"kind": kind}))
    for name in ("$1", "$a-b", "$a b", "a", "$" + "x" * 65):
        assert "error" in parse_line(f'let {name} = analyze {{"kind":"capture"}}')
    assert parse_line("finish when=verified") == {"kind": "verified-finish"}
    for tail in ("finish now", "finish when=verified\nread A1", "finish when=verified\nunknown"):
        records = parse_block("```cmd\nset A1 2\n" + tail + "\n```")["commands"]
        assert records and all("error" in rec for rec in records)
    assert all("error" in rec for rec in parse_block("```cmd\nset A1 2\nfinish when=verif")["commands"])
    assert verified_frame_error("```cmd\nset A1 2\nfinish when=verified") is None
    assert not block_is_complete(parse_block("```cmd\nfinish when=verified\n```"))
    verified = "```cmd\nset S!A1 5\nfinish when=verified\n```"
    plain = "```cmd\nset S!A2 6\n```"
    for response in ("Explanation\n" + verified, verified + "\nExplanation", verified + "\n# comment"):
        assert verified_frame_error(response) is None
    for response in (verified + "\n" + plain, plain + "\n" + verified,
                     verified.replace("set S!A1 5", 'set S!A1 "```"')):
        records = parse_block(response)["commands"]
        assert records and all("error" in rec for rec in records)
        assert verified_frame_error(response)
    assert verified_frame_error(" \n" + verified + "\n ") is None
    assert parse_block("Explanation\n" + plain + "\n" + plain)["commands"] == [{"verb": "set", "cell": "S!A2", "value": "6"}]

    # C1 richer algebra (TS parity: expr-grammar.test.ts / each-grammar.test.ts).
    assert parse_line("let $n = 5") == {"kind": "let", "name": "n", "pipeline": {"kind": "pipeline", "source": {"src": "number", "value": 5}, "stages": []}}
    assert parse_line('let $s = "hi there"')["pipeline"]["source"] == {"src": "text", "value": "hi there"}
    assert "error" in parse_line('let $s = "a" b')
    lit = parse_line('let $t = table [["region","amount"],["East",100],["West",true,"x"],["North"]]')
    assert lit["pipeline"]["source"] == {"src": "table", "columns": ["region", "amount"],
                                         "rows": [["East", "100"], ["West", "true"], ["North", ""]]}
    assert parse_line('table [["a"],["1"]] | count')["stages"] == [{"name": "count", "arg": ""}]
    for bad in ("let $t = table [not json", "let $t = table []", 'let $t = table [["a",""]]',
                'let $t = table [["a","a"]]', 'let $t = table [["a"],[{"x":1}]]'):
        assert "error" in parse_line(bad), bad
    assert "composed" in parse_line("table Sales!A1:B9 | count")["error"]
    for stage in ("join $o on id", "group region agg sum(amount) as total", "distinct region", "rename a=b",
                  "derive u = upper(name)", "filter region=East and not (amount<5 or amount>9)"):
        assert "error" not in parse_line("let $x = $t | " + stage), stage
    assert "unknown transform" in parse_line("let $x = $t | joinx $o")["error"]
    each = parse_block("```cmd\nlet $t = read A1:B4\neach $r in $t:\n  set Summary!$r.cell $r.amount\n  # note\n\n  comment $r.cell \"x\"\nend\ndone\n```")["commands"]
    assert each[1] == {"kind": "each-block", "rowVar": "r", "tableVar": "t",
                       "body": ["set Summary!$r.cell $r.amount", 'comment $r.cell "x"']}, each
    assert each[2] == {"verb": "done"}
    assert block_is_complete(parse_block("```cmd\neach $r in $t:\nset B2 $r.a\nend\ndone\n```"))
    unterminated = parse_block("```cmd\neach $r in $t:\nset B2 $r.a\n```")["commands"]
    assert unterminated == [{"error": "each $r in $t is missing a closing `end`"}], unterminated
    for program, needle in (("each $r in $t:\nend", "empty body"),
                            ("each $r in $t:\neach $s in $u:\nset B2 1\nend\nend", "nested each"),
                            ("each $r in $t:\nread $r.cell\nend", "effect commands"),
                            ("each $r in $t:\nset B2 $row.a\nend", "unknown row variable $row"),
                            ("each $r in $t:\nset B2 $r\nend", "bare $r"),
                            ("each r in $t:\nset B2 1\nend", "malformed each header"),
                            ("each $r in $r:\nset B2 1\nend", "distinct names"),
                            ("set B2 1\nend", "without a matching")):
        recs = parse_block("```cmd\n" + program + "\ndone\n```")["commands"]
        assert any("error" in r and needle in r["error"] for r in recs), (program, recs)
        assert recs[-1] == {"verb": "done"} or any("error" in r for r in recs), program
    print("SELF-TEST OK — fail-closed `done`, unclosed fence, trailing-token guards hold")
    return 0


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        sys.exit(_self_test())
    else:
        print(json.dumps(parse_block(sys.stdin.read()), indent=2))
