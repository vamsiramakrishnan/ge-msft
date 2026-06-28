#!/usr/bin/env python3
"""
Extract short-lived Gemini Enterprise widget credentials from a locally saved DevTools request.

The public Discovery Engine API does not currently expose the same UI-created skill-agent
upload/delete surface used by Gemini Enterprise's web UI. For dev-only skill updates we therefore
call the widget API, but we keep the manual step narrow:

  - paste/export one authenticated `content-discoveryengine.googleapis.com` cURL/HAR request locally
  - extract only Authorization Bearer, configId/widgetConfigId, x-server-token, and resource ids
  - write the Bearer token to a chmod 0600 temp file
  - print shell exports for update_skills.py

This script deliberately does not parse or replay Google session cookies, SAPISIDHASH, or getoxsrf
state. Those belong to the browser session, not to repository automation.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import stat
import sys
from pathlib import Path
from typing import Any

import create_skill


DEFAULT_TOKEN_FILE = Path("/tmp/ge-widget-token")


def _walk_json(value: Any):
    yield value
    if isinstance(value, dict):
        for child in value.values():
            yield from _walk_json(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk_json(child)


def _header_from_har(data: Any, name: str) -> str | None:
    wanted = name.lower()
    for node in _walk_json(data):
        if not isinstance(node, dict):
            continue
        headers = node.get("headers")
        if not isinstance(headers, list):
            continue
        for header in headers:
            if not isinstance(header, dict):
                continue
            if str(header.get("name", "")).lower() == wanted:
                value = header.get("value")
                return str(value) if value is not None else None
    return None


def _first_json_value(data: Any, *names: str) -> str | None:
    wanted = set(names)
    for node in _walk_json(data):
        if not isinstance(node, dict):
            continue
        for name in wanted:
            value = node.get(name)
            if isinstance(value, str) and value.strip():
                return value.strip()
    return None


def _first_nested_json_value(data: Any, *names: str) -> str | None:
    direct = _first_json_value(data, *names)
    if direct:
        return direct
    for node in _walk_json(data):
        if not isinstance(node, str):
            continue
        nested = _try_json(node)
        if nested is None:
            continue
        found = _first_json_value(nested, *names)
        if found:
            return found
    return None


def _text_from_file(path: Path | None) -> str:
    if path is None or str(path) == "-":
        return sys.stdin.read()
    return path.read_text(encoding="utf-8")


def _try_json(text: str) -> Any | None:
    stripped = text.strip()
    if not stripped or stripped[0] not in "[{":
        return None
    try:
        return json.loads(stripped)
    except json.JSONDecodeError:
        return None


def _regex_first(patterns: list[str], text: str) -> str | None:
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
        if match:
            return match.group(1).strip()
    return None


def extract_widget_credentials(text: str) -> dict[str, str]:
    data = _try_json(text)
    out: dict[str, str] = {}

    if data is not None:
        auth = _header_from_har(data, "authorization")
        server_token = _header_from_har(data, "x-server-token")
        config_id = _first_nested_json_value(data, "configId", "widgetConfigId")
        if auth:
            out["token"] = auth.removeprefix("Bearer ").strip()
        if server_token:
            out["server_token"] = server_token.strip()
        if config_id:
            out["config_id"] = config_id.strip()

    out.setdefault(
        "token",
        _regex_first(
            [
                r"authorization:\s*Bearer\s+([A-Za-z0-9_.=-]+)",
                r'"authorization"\s*:\s*"Bearer\s+([A-Za-z0-9_.=-]+)"',
                r"'authorization:\s*Bearer\s+([A-Za-z0-9_.=-]+)'",
                r"-H\s+['\"]authorization:\s*Bearer\s+([A-Za-z0-9_.=-]+)['\"]",
            ],
            text,
        )
        or "",
    )
    out.setdefault(
        "server_token",
        _regex_first(
            [
                r"x-server-token:\s*([^'\"\s\\]+)",
                r'"x-server-token"\s*:\s*"([^"]+)"',
                r"-H\s+['\"]x-server-token:\s*([^'\"]+)['\"]",
            ],
            text,
        )
        or "",
    )
    out.setdefault(
        "config_id",
        _regex_first(
            [
                r'"configId"\s*:\s*"([^"]+)"',
                r'"widgetConfigId"\s*:\s*"([^"]+)"',
            ],
            text,
        )
        or "",
    )

    resource = _regex_first(
        [
            r"projects/(\d+)/locations/([^/]+)/collections/default_collection/engines/([^/\"'?\s]+)",
        ],
        text,
    )
    if resource:
        match = re.search(
            r"projects/(\d+)/locations/([^/]+)/collections/default_collection/engines/([^/\"'?\s]+)",
            text,
        )
        assert match is not None
        out["project_number"] = match.group(1)
        out["location"] = match.group(2)
        out["engine"] = match.group(3)

    return {k: v for k, v in out.items() if v}


def _write_token(path: Path, token: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(token, encoding="utf-8")
    path.chmod(stat.S_IRUSR | stat.S_IWUSR)


def _write_env_file(path: Path, exports: dict[str, str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"export {key}={_shell_quote(value)}\n" for key, value in exports.items()]
    path.write_text("".join(lines), encoding="utf-8")
    path.chmod(stat.S_IRUSR | stat.S_IWUSR)


def _shell_quote(value: str) -> str:
    return "'" + value.replace("'", "'\"'\"'") + "'"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract widget API env vars from a saved Gemini Enterprise cURL/HAR request."
    )
    parser.add_argument("input", nargs="?", type=Path, help="curl/HAR/text file; omit or use - for stdin")
    parser.add_argument("--token-file", type=Path, default=DEFAULT_TOKEN_FILE)
    parser.add_argument("--env-file", type=Path, help="write sourceable shell exports to this file")
    parser.add_argument("--project", default=os.environ.get("GE_PROJECT", ""))
    parser.add_argument("--agent-surface", default=os.environ.get("GE_SURFACE_COMMANDER_AGENT_ID", ""))
    parser.add_argument("--agent-planner", default=os.environ.get("GE_COMMAND_PLANNER_AGENT_ID", ""))
    args = parser.parse_args(argv)

    text = _text_from_file(args.input)
    creds = extract_widget_credentials(text)
    token = creds.get("token", "")
    if not token:
        raise SystemExit(
            "No widget Bearer token found. Save a request to "
            "content-discoveryengine.googleapis.com with an Authorization: Bearer header."
        )
    create_skill._validate_widget_bearer_token(token)
    _write_token(args.token_file, token)

    exports: dict[str, str] = {"GE_WIDGET_BEARER_TOKEN_FILE": str(args.token_file)}
    if args.project:
        exports["GE_PROJECT"] = args.project
    if creds.get("project_number"):
        exports["GE_PROJECT_NUMBER"] = creds["project_number"]
    if creds.get("location"):
        exports["GE_LOCATION"] = creds["location"]
    if creds.get("engine"):
        exports["GE_ENGINE"] = creds["engine"]
    if creds.get("config_id"):
        exports["GE_WIDGET_CONFIG_ID"] = creds["config_id"]
    if creds.get("server_token"):
        exports["GE_WIDGET_SERVER_TOKEN"] = creds["server_token"]
    if args.agent_surface:
        exports["GE_SURFACE_COMMANDER_AGENT_ID"] = args.agent_surface
    if args.agent_planner:
        exports["GE_COMMAND_PLANNER_AGENT_ID"] = args.agent_planner

    print(f"wrote widget bearer token to {args.token_file} (mode 0600)")
    if args.env_file:
        _write_env_file(args.env_file, exports)
        print(f"wrote widget env exports to {args.env_file} (mode 0600)")
    print("\n# Run these exports in your shell:")
    for key, value in exports.items():
        print(f"export {key}={_shell_quote(value)}")
    print("\n# Then run:")
    print("python3 skill/update_skills.py --api-mode widget --replace --yes --live")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
