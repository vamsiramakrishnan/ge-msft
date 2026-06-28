#!/usr/bin/env python3
"""
command_help.py — render targeted help/playbooks from the generated m365 CLI manifest.

The data source is scripts/m365-cli-1.0.json, emitted from @ge/contracts. This keeps
skill-side progressive disclosure in sync with the runtime parser and capability grammar.

Usage:
  python3 scripts/command_help.py
  python3 scripts/command_help.py shape
  python3 scripts/command_help.py --json shape
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from parse_commands import COMMAND_HELP, READ_VERBS, CONTROL_VERBS, WRITE_VERBS  # noqa: E402


def _topic_key(topic: str | None) -> str | None:
    if not topic:
        return None
    t = topic.strip().lower()
    return t[1:] if t.startswith("/") else t


def _section(title: str, lines):
    if not lines:
        return []
    return [f"{title}:"] + [f"  {i + 1}. {line}" for i, line in enumerate(lines)]


def render(topic: str | None = None) -> str:
    key = _topic_key(topic)
    if not key:
        read = ", ".join(sorted(READ_VERBS))
        write = ", ".join(sorted(WRITE_VERBS))
        control = ", ".join(sorted(CONTROL_VERBS))
        return "\n".join(
            [
                "m365 CLI help",
                f"Read verbs: {read}",
                f"Write verbs: {write}",
                f"Control verbs: {control}",
                "Targeted help: python3 scripts/command_help.py <command>",
                "In a live command turn: help <command> or <command> -h",
            ]
        )

    entry = COMMAND_HELP.get(key)
    if not entry:
        return f'No generated help for "{topic}". Run without a topic to list commands.'

    out = [
        f"Command: {entry['command']}",
        f"Use when: {entry['useWhen']}",
        f"Syntax: {entry['syntax']}",
    ]
    out.extend(_section("Discovery sequence", entry.get("discovery", [])))
    out.extend(_section("Action sequence", entry.get("sequence", [])))
    out.extend(_section("Examples", entry.get("examples", [])))
    out.extend(_section("Do not", entry.get("doNot", [])))
    out.extend(_section("Failure modes", entry.get("failureModes", [])))
    out.extend(_section("Safety", entry.get("safety", [])))
    return "\n".join(out)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="command_help", description="render generated m365 CLI help")
    ap.add_argument("topic", nargs="?")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    key = _topic_key(args.topic)
    if args.json:
        payload = COMMAND_HELP if not key else COMMAND_HELP.get(key)
        print(json.dumps(payload, indent=2, sort_keys=True))
    else:
        print(render(args.topic))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
