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
    "set", "suggest", "comment", "format", "reply",
    "slide", "page", "mail", "post", "compose",
    # ADR-0007 host-native Excel kinds — table/chart/cf take a literal range + props
    # (the `format`-style grammar); spill is the table→grid composition sink (`set` widened).
    "table", "chart", "cf", "spill",
}


def _load_language():
    """Load (read, control, write, transforms, effects, kinds, version) from the bundled manifest."""
    try:
        data = json.loads(_MANIFEST_PATH.read_text(encoding="utf-8"))
        return (
            set(data["verbs"]["read"]),
            set(data["verbs"]["control"]),
            set(data["verbs"]["write"]),
            set(data.get("transforms", [])),
            set(data.get("effectVerbs", [])),
            set(data.get("actuationKinds", [])),
            data.get("version", "unknown"),
        )
    except (OSError, KeyError, ValueError):
        # Fallback: the manifest is absent (stripped sandbox) — fall back to HANDLED_WRITE_VERBS so the
        # checker still runs. Parity asserts this matches the manifest, so the fallback can't rot.
        return (
            {"outline", "read", "search"},
            {"done", "help"},
            set(HANDLED_WRITE_VERBS),
            {"filter", "select", "sum", "avg", "min", "max", "count", "sort", "head", "tail"},
            set(HANDLED_WRITE_VERBS),
            set(),
            "fallback",
        )


(
    READ_VERBS,
    CONTROL_VERBS,
    WRITE_VERBS,
    TRANSFORM_NAMES,
    EFFECT_VERBS,
    ACTUATION_KINDS,
    LANGUAGE_VERSION,
) = _load_language()
ALL_VERBS = READ_VERBS | CONTROL_VERBS | WRITE_VERBS

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


def parse_line(line: str):
    """Parse one command line → a dict record, or {'error': '...'} (the corrective contract)."""
    line = line.strip()
    if not line or line.startswith("#"):
        return None
    parts = line.split(None, 1)
    verb = parts[0]
    rest = parts[1].strip() if len(parts) > 1 else ""

    # ADR-0008 §two-tier — a `/<kind>` line is the SPECIALIZED surface (the long-tail catalogue),
    # dispatched before the core-verb check. The command name IS the ActuationKind (drift-free).
    if verb.startswith("/"):
        return _parse_invoke(verb, rest)

    if verb not in ALL_VERBS:
        return {"error": f"unknown verb '{verb}'{_did_you_mean(verb)} (run 'help')"}

    # No-argument verbs. Consume the FULL line: a trailing token is malformed input, not something
    # to silently drop (review Finding #8/#10 — a dropped `done`-block tail hid real parse failures).
    if verb in ("outline", "done", "help"):
        if rest:
            return {"error": f"{verb} takes no arguments — got {rest!r} (usage: {verb})"}
        return {"verb": verb}
    if verb == "read":
        return {"verb": "read", "selector": rest}  # empty ⇒ whole doc
    if verb == "search":
        if not rest:
            return {"error": "search needs text — usage: search <text>"}
        return {"verb": "search", "text": rest.strip('"')}

    if verb == "set":
        sp = re.search(r"\s", rest)
        if not sp:
            return {"error": "set needs a cell and a value — usage: set <A1> <value|=formula>"}
        cell, value = rest[: sp.start()], rest[sp.start():].strip()
        if not cell or not value:
            return {"error": "set needs a cell and a value — usage: set <A1> <value|=formula>"}
        return {"verb": "set", "cell": cell, "value": value}

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

    if failures:
        print("SELF-TEST FAIL", file=sys.stderr)
        for f in failures:
            print(f"  - {f}", file=sys.stderr)
        return 1
    print("SELF-TEST OK — fail-closed `done`, unclosed fence, trailing-token guards hold")
    return 0


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        sys.exit(_self_test())
    else:
        print(json.dumps(parse_block(sys.stdin.read()), indent=2))
