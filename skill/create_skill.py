#!/usr/bin/env python3
"""
Create a Gemini Enterprise skill programmatically. The current Gemini Enterprise web UI uses the
content-discoveryengine widget API for skill creation plus Google resumable upload for zip bundles:

  1) POST /v1alpha/locations/<location>/widgetCreateAgent
       -> returns agent.name, an opaque numeric id such as "8870098647237058037"
  2) POST /upload/v1alpha/{assistant}/agents/<agent.name>/files:upload
       x-goog-upload-command: start
       x-goog-upload-protocol: resumable
  3) POST the zip bytes to the returned upload URL
       x-goog-upload-command: upload, finalize

This is not the same as the public discoveryengine.googleapis.com raw upload path. The raw legacy
path is still available through --api-mode legacy for older/admin environments, but --api-mode
widget is the default because it matches the live UI traffic.

Two methods:

  Method A — single-file skill (instruction only):
      widget mode: widgetCreateAgent with skillAgentDefinition.instruction
      legacy mode: POST {assistant}/agents?agentId=<id>
        { displayName, description, skillAgentDefinition: { instruction: "<full markdown>" } }

  Method B — multi-file bundle (SKILL.md + references/ + scripts/ + assets/):  [default]
      widget mode:
        1) widgetCreateAgent with a placeholder skillAgentDefinition.instruction
        2) resumable upload the zip to agents/<returned agent.name>/files:upload
      legacy mode:
        1) POST {assistant}/agents?agentId=<id> with a placeholder instruction
        2) raw upload the zip to agents/<id>/files:upload
         -> the server unpacks the zip: SKILL.md body -> instruction, the rest -> subfiles.
      3) GET to verify.

This is the same result the UI produces (create -> files:upload -> getAgentView), with a plain
OAuth Bearer token (ADC) instead of SAPISIDHASH/widget config.

Auth:
  * widget mode: GE_AUTH_MODE=widget and GE_WIDGET_BEARER_TOKEN copied from the authenticated
    Gemini Enterprise web session. It is short-lived; do not commit or log it.
  * legacy mode: ADC or GE_AUTH_MODE=gcloud (gcloud auth print-access-token).

SAFETY (review Finding #8):
  * There are NO baked-in project/engine identifiers. Any LIVE operation REQUIRES the GE_PROJECT,
    GE_PROJECT_NUMBER and GE_ENGINE environment variables; the tool refuses (and names the missing
    vars) otherwise — live mode cannot be selected accidentally.
  * DRY-RUN is the default. Pass --live to actually talk to the API. The destructive / tenant-wide
    operations --replace (delete-then-recreate) and --share (scope=ALL_USERS) additionally require
    an explicit --yes, and the exact target project + engine + agent id are printed before they run.
  * Every request carries an explicit timeout and bounded retries, and calls raise_for_status().

Usage:
  python3 create_skill.py                              # DRY-RUN (default): print the plan, no I/O
  python3 create_skill.py --live                       # Method B via widget API
  python3 create_skill.py --live --zip path/skill.zip  # Method B with a specific zip
  python3 create_skill.py --live --single-file SKILL.md  # Method A: inline instruction from markdown
  python3 create_skill.py --live --api-mode legacy --replace --yes  # legacy delete first
  python3 create_skill.py --live --share --yes         # set sharingConfig.scope=ALL_USERS after create
"""

import argparse
import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

import requests
from requests import HTTPError
from requests.adapters import HTTPAdapter

try:  # urllib3 ships with requests; import path differs across versions.
    from urllib3.util.retry import Retry
except ImportError:  # pragma: no cover - very old requests vendored urllib3
    from requests.packages.urllib3.util.retry import Retry  # type: ignore

LOCATION = os.environ.get("GE_LOCATION", "global")
AGENT_ID = os.environ.get("GE_AGENT_ID", "m365-surface-commander")  # slug referenced in skillsSpec
DISPLAY_NAME = "M365 Surface Commander"
DESCRIPTION = (
    "Reads, analyzes, and edits the Microsoft 365 document the user has open via a "
    "compact command-line protocol the Office add-in applies as reviewable changes."
)

API = "https://discoveryengine.googleapis.com/v1alpha"
UPLOAD_API = "https://discoveryengine.googleapis.com/upload/v1alpha"
CONTENT_API = "https://content-discoveryengine.googleapis.com/v1alpha"
CONTENT_UPLOAD_API = "https://content-discoveryengine.googleapis.com/upload/v1alpha"

# Request hardening: every call gets a finite timeout and a small, bounded retry budget.
HTTP_TIMEOUT = 60  # seconds
MAX_RETRIES = 3

REQUIRED_ENV = ("GE_PROJECT", "GE_PROJECT_NUMBER", "GE_ENGINE")
WIDGET_ORIGIN = "https://vertexaisearch.cloud.google"
API_MODES = ("widget", "legacy")


@dataclass(frozen=True)
class LiveConfig:
    """Fully-resolved target for a live operation. Constructed only from explicit env vars."""

    project: str
    project_number: str
    engine: str
    location: str

    @property
    def assistant(self) -> str:
        return (
            f"projects/{self.project_number}/locations/{self.location}"
            f"/collections/default_collection/engines/{self.engine}/assistants/default_assistant"
        )


@dataclass(frozen=True)
class WidgetConfig:
    config_id: str
    server_token: str | None = None


def resolve_live_config() -> LiveConfig:
    """Build a LiveConfig from the environment, or exit clearly if any required var is unset.

    There are deliberately NO defaults for the project/engine identifiers: live mode targets a real
    tenant and must be selected on purpose.
    """
    missing = [name for name in REQUIRED_ENV if not os.environ.get(name)]
    if missing:
        raise SystemExit(
            "Refusing to run a live operation: missing required environment "
            f"variable(s): {', '.join(missing)}.\n"
            "Set GE_PROJECT, GE_PROJECT_NUMBER and GE_ENGINE (and optionally GE_LOCATION, "
            "GE_AGENT_ID) to point at your tenant before using --live."
        )
    return LiveConfig(
        project=os.environ["GE_PROJECT"],
        project_number=os.environ["GE_PROJECT_NUMBER"],
        engine=os.environ["GE_ENGINE"],
        location=os.environ.get("GE_LOCATION", LOCATION),
    )


def resolve_widget_config() -> WidgetConfig:
    config_id = os.environ.get("GE_WIDGET_CONFIG_ID")
    if not config_id:
        raise SystemExit(
            "Widget API mode requires GE_WIDGET_CONFIG_ID (for your dev app this is the "
            "widget config GUID from the Gemini Enterprise URL/captured widget calls)."
        )
    return WidgetConfig(
        config_id=config_id,
        server_token=os.environ.get("GE_WIDGET_SERVER_TOKEN"),
    )


def _widget_bearer_token() -> str:
    token = os.environ.get("GE_WIDGET_BEARER_TOKEN", "").strip()
    token_file = os.environ.get("GE_WIDGET_BEARER_TOKEN_FILE", "").strip()
    if not token and token_file:
        token = Path(token_file).read_text(encoding="utf-8").strip()
    if not token:
        raise SystemExit(
            "Widget API live mode requires GE_WIDGET_BEARER_TOKEN or GE_WIDGET_BEARER_TOKEN_FILE. "
            "Use the short-lived Bearer token from an authenticated Gemini Enterprise web request; "
            "do not commit it or paste it into logs."
        )
    return token.removeprefix("Bearer ").strip()


def session(cfg: LiveConfig, api_mode: str = "legacy") -> requests.Session:
    """Authenticated requests.Session with bounded retries. Imports google.auth lazily so the
    tooling (and its offline tests) can be imported without the dependency installed."""
    s = requests.Session()
    retry = Retry(
        total=MAX_RETRIES,
        backoff_factor=0.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "POST", "PATCH", "DELETE"}),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    if api_mode == "widget" or os.environ.get("GE_AUTH_MODE") == "widget":
        token = _widget_bearer_token()
        s.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Origin": WIDGET_ORIGIN,
                "Referer": f"{WIDGET_ORIGIN}/",
            }
        )
        server_token = os.environ.get("GE_WIDGET_SERVER_TOKEN")
        if server_token:
            s.headers.update({"x-server-token": server_token})
        return s

    if os.environ.get("GE_AUTH_MODE") == "gcloud":
        token = subprocess.check_output(
            ["gcloud", "auth", "print-access-token"], text=True
        ).strip()
    else:
        import google.auth
        import google.auth.transport.requests

        creds, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
        creds.refresh(google.auth.transport.requests.Request())
        token = creds.token
    s.headers.update({"Authorization": f"Bearer {token}", "X-Goog-User-Project": cfg.project})
    return s


def _widget_request(widget: WidgetConfig, key: str, payload: dict) -> dict:
    return {
        "configId": widget.config_id,
        "additionalParams": {"token": "-", "origin": "ORIGIN_UNSPECIFIED"},
        key: payload,
    }


def _response_body_for_error(r: requests.Response) -> str:
    text = r.text.strip()
    if len(text) > 2000:
        return text[:2000] + "...<truncated>"
    return text


def _raise_for_status_with_body(r: requests.Response, op: str) -> None:
    try:
        r.raise_for_status()
    except HTTPError as exc:
        body = _response_body_for_error(r)
        raise HTTPError(f"{op} failed: HTTP {r.status_code}: {body}", response=r) from exc


def _looks_like_missing_agent(r: requests.Response) -> bool:
    if r.status_code == 404:
        return True
    if r.status_code != 400:
        return False
    body = _response_body_for_error(r).lower()
    missing_markers = (
        "not found",
        "not_found",
        "does not exist",
        "no such agent",
        "not exist",
        "already deleted",
    )
    return any(marker in body for marker in missing_markers)


def delete_agent(s: requests.Session, cfg: LiveConfig, agent_id: str):
    r = s.delete(f"{API}/{cfg.assistant}/agents/{agent_id}", timeout=HTTP_TIMEOUT)
    print(f"  delete {agent_id}: HTTP {r.status_code}")
    # 404 is fine (nothing to delete, idempotent); surface every other failure.
    if r.status_code == 404:
        return r
    r.raise_for_status()
    return r


def delete_widget_agent(
    s: requests.Session,
    cfg: LiveConfig,
    widget: WidgetConfig,
    agent_name: str,
    *,
    missing_ok: bool = True,
):
    r = s.post(
        f"{CONTENT_API}/locations/{cfg.location}/widgetDeleteAgent",
        json=_widget_request(widget, "deleteAgentRequest", {"name": agent_name}),
        timeout=HTTP_TIMEOUT,
    )
    print(f"  delete {agent_name}: HTTP {r.status_code}")
    if missing_ok and _looks_like_missing_agent(r):
        if r.status_code != 404:
            print(f"  delete {agent_name}: treating missing/stale agent as already deleted")
        return r
    _raise_for_status_with_body(r, "widgetDeleteAgent")
    return r


def create_shell(
    s: requests.Session,
    cfg: LiveConfig,
    agent_id: str,
    instruction: str,
    display_name: str = DISPLAY_NAME,
    description: str = DESCRIPTION,
) -> dict:
    r = s.post(
        f"{API}/{cfg.assistant}/agents",
        params={"agentId": agent_id},
        json={
            "displayName": display_name,
            "description": description,
            "skillAgentDefinition": {"instruction": instruction},
        },
        timeout=HTTP_TIMEOUT,
    )
    _raise_for_status_with_body(r, "widgetCreateAgent")
    return r.json()


def create_widget_agent(
    s: requests.Session,
    cfg: LiveConfig,
    widget: WidgetConfig,
    instruction: str,
    display_name: str = DISPLAY_NAME,
    description: str = DESCRIPTION,
) -> dict:
    r = s.post(
        f"{CONTENT_API}/locations/{cfg.location}/widgetCreateAgent",
        json=_widget_request(
            widget,
            "createAgentRequest",
            {
                "agent": {
                    "displayName": display_name,
                    "description": description,
                    "skillAgentDefinition": {"instruction": instruction},
                },
                "defaultFilesSkipped": True,
            },
        ),
        timeout=HTTP_TIMEOUT,
    )
    r.raise_for_status()
    d = r.json()
    agent = d.get("agent")
    if not isinstance(agent, dict) or not agent.get("name"):
        raise RuntimeError(f"widgetCreateAgent returned no agent.name: {d!r}")
    return agent


def upload_zip(s: requests.Session, cfg: LiveConfig, agent_id: str, zip_path: Path) -> dict:
    r = s.post(
        f"{UPLOAD_API}/{cfg.assistant}/agents/{agent_id}/files:upload",
        params={"upload_protocol": "raw"},
        headers={"Content-Type": "application/zip"},
        data=zip_path.read_bytes(),
        timeout=HTTP_TIMEOUT,
    )
    r.raise_for_status()
    return r.json()


def upload_zip_resumable(
    s: requests.Session, cfg: LiveConfig, agent_name: str, zip_path: Path
) -> dict:
    data = zip_path.read_bytes()
    start = s.post(
        f"{CONTENT_UPLOAD_API}/{cfg.assistant}/agents/{agent_name}/files:upload",
        headers={
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            "x-goog-upload-command": "start",
            "x-goog-upload-file-name": zip_path.name,
            "x-goog-upload-header-content-length": str(len(data)),
            "x-goog-upload-protocol": "resumable",
        },
        data=b"",
        timeout=HTTP_TIMEOUT,
    )
    _raise_for_status_with_body(start, "files:upload start")
    upload_url = start.headers.get("x-goog-upload-url")
    if not upload_url:
        raise RuntimeError(
            "resumable upload start returned no x-goog-upload-url header; "
            f"status={start.status_code}"
        )

    final = s.post(
        upload_url,
        headers={
            "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
            "x-goog-upload-command": "upload, finalize",
            "x-goog-upload-file-name": zip_path.name,
            "x-goog-upload-offset": "0",
        },
        data=data,
        timeout=HTTP_TIMEOUT,
    )
    _raise_for_status_with_body(final, "files:upload finalize")
    if final.content:
        try:
            return final.json()
        except ValueError:
            return {"raw": final.text}
    return {}


def share(s: requests.Session, cfg: LiveConfig, agent_id: str):
    r = s.patch(
        f"{API}/{cfg.assistant}/agents/{agent_id}",
        params={"updateMask": "sharingConfig"},
        json={"sharingConfig": {"scope": "ALL_USERS"}},
        timeout=HTTP_TIMEOUT,
    )
    print(f"  share {agent_id}: HTTP {r.status_code}")
    _raise_for_status_with_body(r, "widgetGetAgentView")
    return r


def get_widget_agent_view(
    s: requests.Session, cfg: LiveConfig, widget: WidgetConfig, agent_name: str
) -> dict:
    r = s.post(
        f"{CONTENT_API}/locations/{cfg.location}/widgetGetAgentView",
        json=_widget_request(widget, "getAgentViewRequest", {"name": agent_name}),
        timeout=HTTP_TIMEOUT,
    )
    r.raise_for_status()
    return r.json()


def _extract_widget_agent(view: dict) -> dict:
    candidates = [
        view.get("agent"),
        view.get("agentView"),
        view.get("agentView", {}).get("agent") if isinstance(view.get("agentView"), dict) else None,
    ]
    for candidate in candidates:
        if isinstance(candidate, dict) and candidate.get("name"):
            return candidate
    return view


def show_widget(s: requests.Session, cfg: LiveConfig, widget: WidgetConfig, agent_name: str) -> None:
    d = get_widget_agent_view(s, cfg, widget, agent_name)
    agent = _extract_widget_agent(d)
    sd = agent.get("skillAgentDefinition", {}) if isinstance(agent, dict) else {}
    print(f"  name:        {agent.get('name', agent_name)}")
    print(f"  displayName: {agent.get('displayName')}")
    print(f"  state:       {agent.get('state')}")
    print(f"  instruction: {len(sd.get('instruction',''))} chars")
    print(f"  subfiles:    {[f.get('fileName') for f in sd.get('subfiles', [])] or '(none)'}")
    print("\n  reference in widgetStreamAssist via:")
    print(
        "    "
        + f'"skillsSpec": {{"skills": [{{"name": "{cfg.assistant}/agents/{agent.get("name", agent_name)}"}}]}}'
    )


def show(s: requests.Session, cfg: LiveConfig, agent_id: str) -> None:
    r = s.get(f"{API}/{cfg.assistant}/agents/{agent_id}", timeout=HTTP_TIMEOUT)
    r.raise_for_status()
    d = r.json()
    sd = d.get("skillAgentDefinition", {})
    print(f"  name:        {d.get('name','').split('/')[-1]}")
    print(f"  displayName: {d.get('displayName')}")
    print(f"  state:       {d.get('state')}")
    print(f"  instruction: {len(sd.get('instruction',''))} chars")
    print(f"  subfiles:    {[f['fileName'] for f in sd.get('subfiles',[])] or '(none)'}")
    print("\n  reference in streamAssist via:")
    print(f'    "skillsSpec": {{"skills": [{{"name": "{d.get("name")}"}}]}}')


def _print_target(cfg: LiveConfig, agent_id: str, op: str) -> None:
    """Print the exact destructive target before acting on it."""
    print(f"  >>> {op} TARGET")
    print(f"      project:        {cfg.project}")
    print(f"      project_number: {cfg.project_number}")
    print(f"      engine:         {cfg.engine}")
    print(f"      location:       {cfg.location}")
    print(f"      agent_id:       {agent_id}")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Provision the m365-surface-commander GE skill.")
    ap.add_argument("--zip", default=str(Path(__file__).parent / "m365-surface-commander.zip"))
    ap.add_argument(
        "--single-file",
        metavar="SKILL.md",
        help="Method A: create with this file's contents as the inline instruction",
    )
    ap.add_argument("--agent-id", default=AGENT_ID)
    ap.add_argument("--display-name", default=DISPLAY_NAME)
    ap.add_argument("--description", default=DESCRIPTION)
    ap.add_argument(
        "--api-mode",
        choices=API_MODES,
        default="widget",
        help="widget matches the Gemini Enterprise web UI; legacy uses the older public API path",
    )
    ap.add_argument(
        "--upload-existing",
        action="store_true",
        help="widget mode only: skip create and upload the zip to --agent-id / numeric agent name",
    )
    ap.add_argument(
        "--live",
        action="store_true",
        help="actually talk to the API (default is a dry-run that prints the plan and exits)",
    )
    ap.add_argument(
        "--replace",
        action="store_true",
        help="DESTRUCTIVE: delete an existing agent of this id first (requires --yes)",
    )
    ap.add_argument(
        "--share",
        action="store_true",
        help="TENANT-WIDE: set sharingConfig.scope=ALL_USERS (requires --yes)",
    )
    ap.add_argument(
        "--yes",
        action="store_true",
        help="confirm destructive/tenant-wide operations (--replace / --share)",
    )
    args = ap.parse_args(argv)

    destructive = []
    if args.replace:
        destructive.append("--replace")
    if args.share:
        destructive.append("--share")
    if destructive and not args.yes:
        raise SystemExit(
            f"Refusing {' and '.join(destructive)} without confirmation. "
            "These are destructive / tenant-wide; re-run with --yes to proceed."
        )
    if args.api_mode == "widget" and args.replace:
        raise SystemExit("--replace is only implemented for --api-mode legacy")
    if args.api_mode == "widget" and args.share:
        raise SystemExit("--share is only implemented for --api-mode legacy")
    if args.upload_existing and args.api_mode != "widget":
        raise SystemExit("--upload-existing is only valid with --api-mode widget")
    if args.upload_existing and args.single_file:
        raise SystemExit("--upload-existing only applies to zip bundle upload")

    # Resolve the live target up front so even a dry-run shows the real (env-provided) plan and a
    # live run cannot pick up accidental defaults.
    cfg = resolve_live_config()
    widget = resolve_widget_config() if args.api_mode == "widget" else None

    method = "A (single-file)" if args.single_file else "B (bundle upload)"
    print(f"Plan: provision agent '{args.agent_id}' via Method {method} ({args.api_mode} API)")
    _print_target(cfg, args.agent_id, "PROVISION")
    if widget:
        print(f"      widget_config:  {widget.config_id}")
        if args.upload_existing:
            print("  step: --upload-existing -> skip create and upload to this numeric agent name")
    if args.replace:
        print("  step: --replace -> delete existing agent first")
    if args.share:
        print("  step: --share -> sharingConfig.scope=ALL_USERS (visible to ALL users in tenant)")

    if not args.live:
        print("\nDRY-RUN (default): no API calls made. Re-run with --live to execute.")
        return 0

    s = session(cfg, args.api_mode)

    if args.api_mode == "widget":
        assert widget is not None
        if args.single_file:
            print(f"Method A — widget single-file create from {args.single_file}")
            instruction = Path(args.single_file).read_text(encoding="utf-8")
            agent = create_widget_agent(
                s, cfg, widget, instruction, args.display_name, args.description
            )
            agent_name = agent["name"]
        else:
            zip_path = Path(args.zip)
            if not zip_path.exists():
                raise SystemExit(f"zip not found: {zip_path}")
            print(f"Method B — widget bundle upload from {zip_path}")
            if args.upload_existing:
                agent_name = args.agent_id
                print(f"  1) use existing agent name {agent_name}")
            else:
                print("  1) create widget shell agent")
                agent = create_widget_agent(
                    s,
                    cfg,
                    widget,
                    "placeholder — replaced by SKILL.md on upload",
                    args.display_name,
                    args.description,
                )
                agent_name = agent["name"]
                print(f"     created agent.name {agent_name}")
            print("  2) resumable upload zip (server unpacks)")
            upload_zip_resumable(s, cfg, agent_name, zip_path)
        print("  3) verify")
        show_widget(s, cfg, widget, agent_name)
        return 0

    if args.replace:
        _print_target(cfg, args.agent_id, "DELETE")
        delete_agent(s, cfg, args.agent_id)

    if args.single_file:
        print(f"Method A — single-file create from {args.single_file}")
        instruction = Path(args.single_file).read_text(encoding="utf-8")
        create_shell(s, cfg, args.agent_id, instruction, args.display_name, args.description)
    else:
        zip_path = Path(args.zip)
        if not zip_path.exists():
            raise SystemExit(f"zip not found: {zip_path}")
        print(f"Method B — bundle upload from {zip_path}")
        print("  1) create shell agent")
        create_shell(
            s,
            cfg,
            args.agent_id,
            "placeholder — replaced by SKILL.md on upload",
            args.display_name,
            args.description,
        )
        print("  2) upload zip (server unpacks)")
        upload_zip(s, cfg, args.agent_id, zip_path)

    if args.share:
        _print_target(cfg, args.agent_id, "SHARE (ALL_USERS)")
        share(s, cfg, args.agent_id)
    print("  3) verify")
    show(s, cfg, args.agent_id)
    return 0


if __name__ == "__main__":
    sys.exit(main())
