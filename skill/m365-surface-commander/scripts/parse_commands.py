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

READ_VERBS = {"outline", "read", "search"}
CONTROL_VERBS = {"done", "help"}
WRITE_VERBS = {
    "set", "suggest", "comment", "format", "reply",
    "slide", "page", "mail", "post", "compose",
}
ALL_VERBS = READ_VERBS | CONTROL_VERBS | WRITE_VERBS

_FENCE = re.compile(r"```cmd[^\S\n]*\r?\n([\s\S]*?)```", re.IGNORECASE)
# Lenient fallback: an opening ```cmd fence the model never closed — take to end of message.
_FENCE_OPEN = re.compile(r"```cmd[^\S\n]*\r?\n([\s\S]*)$", re.IGNORECASE)
_QUOTED = re.compile(r'"((?:[^"\\]|\\.)*)"')


def extract_command_block(text: str):
    """Return the inner text of the first ```cmd fence, or None (→ re-prompt, not an error).

    Tolerates an unclosed fence: if the model opened ```cmd but never emitted the closing ```,
    we take everything after the opener (a frequent real-world failure mode worth surviving).
    """
    m = _FENCE.search(text)
    if m:
        return m.group(1).strip()
    m = _FENCE_OPEN.search(text)
    if m:
        # drop a dangling trailing fence if present
        return re.sub(r"```\s*$", "", m.group(1)).strip()
    return None


def _scan_quoted(rest: str):
    m = _QUOTED.match(rest.strip())
    if not m:
        return None
    val = m.group(1).replace('\\"', '"').replace("\\\\", "\\")
    return val, rest.strip()[m.end():].strip()


def _did_you_mean(verb: str):
    import difflib
    near = difflib.get_close_matches(verb, sorted(ALL_VERBS), n=1)
    return f" — did you mean '{near[0]}'?" if near else ""


def parse_line(line: str):
    """Parse one command line → a dict record, or {'error': '...'} (the corrective contract)."""
    line = line.strip()
    if not line or line.startswith("#"):
        return None
    parts = line.split(None, 1)
    verb = parts[0]
    rest = parts[1].strip() if len(parts) > 1 else ""

    if verb not in ALL_VERBS:
        return {"error": f"unknown verb '{verb}'{_did_you_mean(verb)} (run 'help')"}

    if verb == "outline":
        return {"verb": "outline"}
    if verb in ("done", "help"):
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
        tail = re.sub(r"^(=>|->)\s*", "", tail.strip())
        second = _scan_quoted(tail)
        if not second:
            return {"error": 'suggest needs two quoted strings — usage: suggest "old" => "new"'}
        return {"verb": "suggest", "oldText": old, "newText": second[0]}

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
        m = re.match(r"(\S+)\s+(.*)", rest)
        if not m:
            return {"error": "format needs a range and at least one key=value — usage: format <range> k=v ..."}
        rng, props_s = m.group(1), m.group(2)
        props = dict(p.split("=", 1) for p in props_s.split() if "=" in p)
        if not props:
            return {"error": f'format expects key=value pairs — got "{props_s}"'}
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

    return {"error": f"unhandled verb '{verb}'"}


def parse_block(model_text: str):
    inner = extract_command_block(model_text)
    if inner is None:
        return {"block": None, "commands": [], "note": "no ```cmd fence (re-prompt, not an error)"}
    out = []
    for line in inner.splitlines():
        rec = parse_line(line)
        if rec is not None:
            out.append(rec)
    return {"block": inner, "commands": out}


def _self_test():
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
    print(json.dumps(parse_block(sample), indent=2))


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        _self_test()
    else:
        print(json.dumps(parse_block(sys.stdin.read()), indent=2))
