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

# ADR-0008 §4 — the verb/transform/kind tables are LOADED from the generated language manifest
# (`m365-cli-<v>.json`, emitted from @ge/contracts), NOT hand-maintained here. This is the
# anti-drift hinge: the TS grammar → the manifest → these tables, so the Python preflight can never
# diverge from the runtime parser. The hardcoded fallback below only runs if the bundled manifest is
# missing (a stripped sandbox); the parity test asserts fallback == manifest, so it cannot drift.
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

CONTEXT_HINTS = {
    "incremental",
    "inline-preferred",
    "reference-preferred",
    "upload-preferred",
    "code-execution-preferred",
    "analytical",
    "full-scope",
}

CONTEXT_KINDS = {
    "selection",
    "document",
    "paragraph",
    "table",
    "range",
    "sheet",
    "slide",
    "shape",
    "image",
    "comment",
    "mail-item",
    "mail-thread",
    "attachment",
    "calendar-event",
    "transcript",
    "page",
    "person",
    "indexed-document",
    "drive-document",
    "file",
    "brief",
}


def _load_language():
    """Load (read, workspace, control, write, transforms, effects, kinds, version) from the bundled manifest."""
    try:
        data = json.loads(_MANIFEST_PATH.read_text(encoding="utf-8"))
        return (
            set(data["verbs"]["read"]),
            set(data["verbs"].get("workspace", [])),
            set(data["verbs"]["control"]),
            set(data["verbs"]["write"]),
            set(data.get("transforms", [])),
            set(data.get("effectVerbs", [])),
            set(data.get("actuationKinds", [])),
            data.get("commandHelp", {}),
            data.get("capabilityRegistry", []),
            data.get("version", "unknown"),
        )
    except (OSError, KeyError, ValueError):
        # Fallback: the manifest is absent (stripped sandbox) — fall back to HANDLED_WRITE_VERBS so the
        # checker still runs. Parity asserts this matches the manifest, so the fallback can't rot.
        return (
            {
                "outline", "read", "search", "list", "inspect", "properties", "comments",
                "attachments", "tables", "slides", "neighbors", "context", "open",
            },
            {"workspace", "save", "cat", "grep", "cp", "mv", "rm", "share"},
            {"done", "help"},
            set(HANDLED_WRITE_VERBS),
            {"filter", "select", "sum", "avg", "min", "max", "count", "sort", "head", "tail"},
            set(HANDLED_WRITE_VERBS),
            set(),
            {},
            [],
            "fallback",
        )


(
    READ_VERBS,
    WORKSPACE_VERBS,
    CONTROL_VERBS,
    WRITE_VERBS,
    TRANSFORM_NAMES,
    EFFECT_VERBS,
    ACTUATION_KINDS,
    COMMAND_HELP,
    CAPABILITY_REGISTRY,
    LANGUAGE_VERSION,
) = _load_language()
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
    val = m.group(1).replace('\\"', '"').replace("\\\\", "\\")
    return val, rest.strip()[m.end():].strip()


def _split_first_arg(rest: str):
    """Split the first argument while preserving single-quoted sheet names.

    Excel selectors such as `'Daily schedule'!C5:I23` must stay one token even though the
    sheet name contains spaces.
    """
    s = rest.lstrip()
    if not s:
        return None
    m = re.match(r'"([^"]*)"(\S*)|\'([^\']*)\'(\S*)|(\S+)', s)
    if not m:
        return None
    if m.group(1) is not None:
        arg = f"{m.group(1)}{m.group(2) or ''}"
    elif m.group(3) is not None:
        arg = f"'{m.group(3)}'{m.group(4) or ''}"
    else:
        arg = m.group(5)
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
_TOKENIZE = re.compile(r'(\w[\w-]*)="([^"]*)"|(\w[\w-]*)=(\S+)|"([^"]*)"|(\S+)')


def _tokenize_args(rest: str):
    positional = []
    props = {}
    for m in _TOKENIZE.finditer(rest):
        if m.group(1) is not None:
            props[m.group(1)] = m.group(2)
        elif m.group(3) is not None:
            props[m.group(3)] = m.group(4)
        elif m.group(5) is not None:
            positional.append(m.group(5))
        else:
            positional.append(m.group(6))
    return positional, props


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


def _did_you_mean(verb: str):
    import difflib
    near = difflib.get_close_matches(verb, sorted(ALL_VERBS), n=1)
    return f" — did you mean '{near[0]}'?" if near else ""


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


def _parse_pipeline_source(head: str):
    if head == "outline":
        return {"src": "outline"}
    if head.startswith("$"):
        if re.match(r"^\$[A-Za-z_][A-Za-z0-9_]*$", head):
            return {"src": "var", "name": head[1:]}
        return {"error": f"invalid pipeline variable '{head}'"}
    parts = head.split(None, 1)
    if len(parts) != 2:
        return {"error": "pipeline needs a source — e.g. read <selector> | <transform>"}
    src, arg = parts[0].lower(), parts[1].strip()
    if src == "read":
        return {"src": "read", "selector": arg}
    if src == "search":
        if not arg:
            return {"error": "search pipeline source needs text"}
        return {"src": "search", "text": arg.strip('"')}
    return {"error": f'unknown pipeline source "{parts[0]}" — use read <selector>, search <text>, outline, or $var'}


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
    if ACTUATION_KINDS and kind not in ACTUATION_KINDS:
        near = difflib.get_close_matches(kind, sorted(ACTUATION_KINDS), n=1)
        tail = f" — did you mean '/{near[0]}'?" if near else ""
        return {"error": f"unknown capability '/{kind}'{tail} (the / surface names an ActuationKind)"}
    positional, props = _tokenize_args(rest)
    return {"verb": "invoke", "kind": kind, "props": props, "args": positional}


def _reject_json_constant(_value):
    raise ValueError("Non-JSON numeric constant")


def parse_line(line: str):
    """Parse one command line → a dict record, or {'error': '...'} (the corrective contract)."""
    line = line.strip()
    if not line or line.startswith("#"):
        return None
    parts = line.split(None, 1)
    raw_verb = parts[0]
    verb = raw_verb.lower()
    rest = parts[1].strip() if len(parts) > 1 else ""

    if verb == "let":
        return _parse_let(line)
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
    if verb in ("outline", "done"):
        if rest:
            return {"error": f"{verb} takes no arguments — got {rest!r} (usage: {verb})"}
        return {"verb": verb}
    if verb == "help":
        return {"verb": "help", **({"topic": rest} if rest else {})}
    if verb == "read":
        return {"verb": "read", "selector": rest}  # empty ⇒ whole doc
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
            return {"error": "cat needs a workspace artifact ref — usage: cat <name|ws:id> [head=N]"}
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
            return {"error": 'grep needs an artifact ref and pattern — usage: grep <name|ws:id> "pattern" [context=N]'}
        ref, tail = split[0], split[1].strip()
        if not tail:
            return {"error": 'grep needs an artifact ref and pattern — usage: grep <name|ws:id> "pattern" [context=N]'}
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
            return {"error": "set needs a cell and a value — usage: set <A1> <value|=formula>"}
        cell, value = rest[: sp.start()], rest[sp.start():].strip()
        if not cell or not value:
            return {"error": "set needs a cell and a value — usage: set <A1> <value|=formula>"}
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
            return {"error": 'reply needs a comment id and a quoted reply — usage: reply <id> "text"'}
        return {"verb": "reply", "commentId": m.group(1), "text": q.group(1)}

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
        return {"verb": verb, slot: q[0]}

    if verb == "slide":
        qs = _QUOTED.findall(rest)
        if not qs:
            return {"error": 'slide needs a quoted title — usage: slide "Title" "bullet" ...'}
        return {"verb": "slide", "title": qs[0], "bullets": qs[1:]}

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
                 "<column|bar|line|pie|scatter|area> <range> [title=\"…\"] [series=rows|columns]")
        positional, props = _tokenize_args(rest)
        if len(positional) < 2:
            return {"error": usage}
        return {"verb": "chart", "chartType": positional[0].lower(),
                "range": positional[1], "props": props}

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
    for line in inner.splitlines():
        rec = parse_line(line)
        if rec is not None:
            out.append(rec)
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

    # A clean, closed, done-bearing block IS complete.
    if not block_is_complete(parse_block("```cmd\nread Sales!C2:C7\ndone\n```")):
        failures.append("clean closed `done` block was NOT reported complete")

    # A no-arg verb with a trailing token is a reported error, not a silent drop.
    if "error" not in (parse_line("done now please") or {}):
        failures.append("`done now please` did not error on its trailing tokens")

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
    if parse_line('open "Sales!A1:C9"') != {"verb": "open", "selector": "Sales!A1:C9"}:
        failures.append("open selector did not parse")
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
    print("SELF-TEST OK — fail-closed `done`, unclosed fence, trailing-token guards hold")
    return 0


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        sys.exit(_self_test())
    else:
        print(json.dumps(parse_block(sys.stdin.read()), indent=2))
