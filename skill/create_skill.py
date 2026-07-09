#!/usr/bin/env python3
"""
Create a Gemini Enterprise skill programmatically. The public Discovery Engine RPC reference lists
WidgetService as an internal, breaking-change-prone widget surface. The Gemini Enterprise web UI
uses that same content-discoveryengine widget surface for private skill management plus Google
resumable upload for zip bundles:

  1) POST /v1alpha/locations/<location>/widgetCreateAgent
       -> returns agent.name, an opaque numeric id such as "8870098647237058037"
  2) POST /upload/v1alpha/{assistant}/agents/<agent.name>/files:upload
       x-goog-upload-command: start
       x-goog-upload-protocol: resumable
  3) POST the zip bytes to the returned upload URL
       x-goog-upload-command: upload, finalize

This is not the same as the public discoveryengine.googleapis.com API path. The documented public
agent API is available through --api-mode public. The old name --api-mode legacy remains as a
backwards-compatible alias because earlier versions of this script used that name for the same
OAuth-based path. --api-mode widget is the default because it matches the live UI traffic.

Two methods:

  Method A — single-file skill (instruction only):
      widget mode: widgetCreateAgent with skillAgentDefinition.instruction
      public mode: POST {assistant}/agents?agentId=<id>
        { displayName, description, skillAgentDefinition: { instruction: "<full markdown>" } }

  Method B — multi-file bundle (SKILL.md + references/ + scripts/ + assets/):  [default]
      widget mode:
        1) widgetCreateAgent with a placeholder skillAgentDefinition.instruction
        2) resumable upload the zip to agents/<returned agent.name>/files:upload
      public mode:
        1) POST {assistant}/agents?agentId=<id> with a placeholder instruction
        2) raw upload the zip to agents/<id>/files:upload
         -> the server unpacks the zip: SKILL.md body -> instruction, the rest -> subfiles.
      3) GET to verify.

This is the same result the UI produces (create -> files:upload -> getAgentView), with the
short-lived widget Bearer token copied from an authenticated Gemini Enterprise request. It is not
SAPISIDHASH/cookie automation and it is not ADC.

Auth:
  * widget mode: GE_AUTH_MODE=widget and GE_WIDGET_BEARER_TOKEN copied from the authenticated
    Gemini Enterprise web session. It is short-lived; do not commit or log it.
  * public mode: ADC or GE_AUTH_MODE=gcloud (gcloud auth print-access-token).

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
  python3 create_skill.py --live --list                # list visible widget skill agents
  python3 create_skill.py --live --api-mode public --replace --yes  # public API delete first
  python3 create_skill.py --live --share --yes         # set sharingConfig.scope=ALL_USERS after create
"""

import argparse
import base64
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import time
import zipfile
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
API_MODES = ("widget", "public", "legacy")
PUBLIC_API_MODES = ("public", "legacy")
WIDGET_AGENT_ORIGINS = ("USER", "GOOGLE")
WIDGET_SKILL_AGENT_FILTER = "agent_type = SKILL_AGENT"


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
    token = token.removeprefix("Bearer ").strip()
    _validate_widget_bearer_token(token)
    return token


def _jwt_payload(token: str) -> dict:
    parts = token.split(".")
    if len(parts) != 3:
        raise ValueError("token is not a JWT")
    payload = parts[1]
    padded = payload + "=" * (-len(payload) % 4)
    raw = base64.urlsafe_b64decode(padded.encode("ascii"))
    parsed = json.loads(raw.decode("utf-8"))
    if not isinstance(parsed, dict):
        raise ValueError("JWT payload is not a JSON object")
    return parsed


def _validate_widget_bearer_token(token: str) -> None:
    try:
        payload = _jwt_payload(token)
    except Exception as exc:
        raise SystemExit(
            "GE_WIDGET_BEARER_TOKEN is not a Vertex AI Search widget JWT. "
            "Do not use gcloud/ADC access tokens here; copy the Bearer token from an authenticated "
            "content-discoveryengine.googleapis.com widget request in the Gemini Enterprise web UI."
        ) from exc

    issuer = payload.get("iss")
    audience = payload.get("aud")
    if issuer != WIDGET_ORIGIN or audience != "https://content-discoveryengine.googleapis.com":
        raise SystemExit(
            "GE_WIDGET_BEARER_TOKEN has the wrong issuer/audience for the widget API. "
            f"Expected iss={WIDGET_ORIGIN!r} and aud='https://content-discoveryengine.googleapis.com'. "
            "Use the Bearer token from a content-discoveryengine.googleapis.com widget request, "
            "not gcloud auth print-access-token or Application Default Credentials."
        )

    exp = payload.get("exp")
    if isinstance(exp, (int, float)) and exp <= time.time():
        raise SystemExit(
            "GE_WIDGET_BEARER_TOKEN is expired. Refresh Gemini Enterprise in the browser, copy a "
            "fresh Bearer token from a content-discoveryengine.googleapis.com request, and rerun."
        )


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


def _api_mode_label(api_mode: str) -> str:
    if api_mode == "legacy":
        return "public (legacy alias)"
    return api_mode


def _widget_request(widget: WidgetConfig, key: str, payload: dict) -> dict:
    return {
        "configId": widget.config_id,
        "additionalParams": {"token": "-", "origin": "ORIGIN_UNSPECIFIED"},
        key: payload,
    }


def widget_skill_resource_name(cfg: LiveConfig, agent_name: str) -> str:
    """Full resource name to put in streamAssist `skillsSpec.skills[].name`."""
    return f"{cfg.assistant}/agents/{agent_name}"


def widget_skill_mention(label: str, agent_name: str) -> str:
    """Widget mention marker to prepend to the query text for deterministic skill routing."""
    return f"[{label}](mention://?uri={agent_name})"


def widget_skill_env_value(label: str, cfg: LiveConfig, agent_name: str) -> str:
    """VITE_GE_*_SKILL value: label plus full skill resource name."""
    return f"{label}={widget_skill_resource_name(cfg, agent_name)}"


def _yaml_single_quote(value: object) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def _frontmatter_version_from_text(text: str) -> str | None:
    source = text
    parts = text.split("---", 2)
    if len(parts) >= 3 and not parts[0].strip():
        source = parts[1]
    for line in source.splitlines():
        stripped = line.strip()
        if stripped.startswith("version:"):
            return stripped.split(":", 1)[1].strip().strip("'\"") or None
    return None


def _stamp_skill_frontmatter(text: str, metadata: dict[str, object]) -> str:
    lines = [
        "x-ge-msft-upload:",
        *[f"  {key}: {_yaml_single_quote(value)}" for key, value in metadata.items() if value],
    ]
    block = "\n".join(lines)
    parts = text.split("---", 2)
    if len(parts) >= 3 and not parts[0].strip():
        frontmatter = parts[1].rstrip()
        body = parts[2].lstrip("\n")
        return f"---\n{frontmatter}\n{block}\n---\n{body}"
    return f"---\n{block}\n---\n{text}"


def stamped_skill_zip(
    zip_path: Path,
    cfg: LiveConfig,
    agent_name: str,
    *,
    label: str | None = None,
) -> Path:
    """Return a temporary zip whose SKILL.md frontmatter records upload provenance.

    The source bundle remains untouched. The uploaded bundle hash is intentionally computed by the
    caller after stamping; putting that final hash into the zip would make the artifact recursive.
    """
    source_bytes = zip_path.read_bytes()
    source_sha = hashlib.sha256(source_bytes).hexdigest()
    skill_label = label or zip_path.stem
    resource = widget_skill_resource_name(cfg, agent_name)
    build_id = f"{skill_label}@{agent_name}+{source_sha[:12]}"

    fd, tmp_name = tempfile.mkstemp(prefix=f"{zip_path.stem}.{agent_name}.", suffix=".zip")
    os.close(fd)
    out_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(zip_path, "r") as src, zipfile.ZipFile(out_path, "w") as dst:
            names = set(src.namelist())
            if "SKILL.md" not in names:
                raise RuntimeError(f"{zip_path} does not contain SKILL.md")
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "SKILL.md":
                    skill_md = data.decode("utf-8")
                    version = _frontmatter_version_from_text(skill_md)
                    data = _stamp_skill_frontmatter(
                        skill_md,
                        {
                            "buildId": build_id,
                            "agentId": agent_name,
                            "resource": resource,
                            "sourceZipSha256": source_sha,
                            "sourceVersion": version or "",
                        },
                    ).encode("utf-8")
                stamped_info = zipfile.ZipInfo(info.filename, info.date_time)
                stamped_info.comment = info.comment
                stamped_info.extra = info.extra
                stamped_info.internal_attr = info.internal_attr
                stamped_info.external_attr = info.external_attr
                stamped_info.compress_type = info.compress_type
                dst.writestr(stamped_info, data)
    except Exception:
        try:
            out_path.unlink(missing_ok=True)
        finally:
            raise
    return out_path


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


def list_agents(s: requests.Session, cfg: LiveConfig) -> dict:
    r = s.get(f"{API}/{cfg.assistant}/agents", timeout=HTTP_TIMEOUT)
    _raise_for_status_with_body(r, "agents.list")
    return r.json()


def delete_widget_agent(
    s: requests.Session,
    cfg: LiveConfig,
    widget: WidgetConfig,
    agent_name: str,
    *,
    missing_ok: bool = True,
):
    url = f"{CONTENT_API}/locations/{cfg.location}/widgetDeleteAgent"
    payload = {
        "configId": widget.config_id,
        "additionalParams": {"token": "-", "origin": "ORIGIN_UNSPECIFIED"},
        "name": agent_name,
    }
    r = s.post(url, json=payload, timeout=HTTP_TIMEOUT)
    print(f"  delete {agent_name}: HTTP {r.status_code}")
    if missing_ok and _looks_like_missing_agent(r):
        if r.status_code != 404:
            print(f"  delete {agent_name}: treating missing/stale agent as already deleted")
        return r
    _raise_for_status_with_body(r, "widgetDeleteAgent")
    return r


def list_widget_agent_views(
    s: requests.Session,
    cfg: LiveConfig,
    widget: WidgetConfig,
    *,
    agent_origin: str = "USER",
    page_size: int = 200,
) -> list[dict]:
    """List visible widget skill agents for the signed-in widget user."""
    if agent_origin not in WIDGET_AGENT_ORIGINS:
        raise ValueError(f"agent_origin must be one of {WIDGET_AGENT_ORIGINS}")
    url = f"{CONTENT_API}/locations/{cfg.location}/widgetListAvailableAgentViews"
    views: list[dict] = []
    page_token: str | None = None
    for _ in range(10):
        request = {
            "pageSize": page_size,
            "filter": WIDGET_SKILL_AGENT_FILTER,
            "agentOrigin": agent_origin,
            **({"pageToken": page_token} if page_token else {}),
        }
        r = s.post(
            url,
            json=_widget_request(widget, "listAvailableAgentViewsRequest", request),
            timeout=HTTP_TIMEOUT,
        )
        _raise_for_status_with_body(r, "widgetListAvailableAgentViews")
        payload = r.json()
        response = payload.get("listAvailableAgentViewsResponse", payload)
        page_views = (
            response.get("agentViews")
            or response.get("availableAgentViews")
            or response.get("agents")
            or []
        )
        if not isinstance(page_views, list):
            raise RuntimeError(f"widgetListAvailableAgentViews returned malformed views: {payload!r}")
        views.extend(view for view in page_views if isinstance(view, dict))
        page_token = response.get("nextPageToken")
        if not page_token:
            break
    return views


def _print_widget_agent_view(cfg: LiveConfig, view: dict) -> None:
    nested_agent = view.get("agent")
    nested_agent = nested_agent if isinstance(nested_agent, dict) else {}
    permissions = view.get("userPermissions")
    permissions = permissions if isinstance(permissions, dict) else {}
    name = str(view.get("name") or nested_agent.get("name") or "(unnamed)")
    display = view.get("displayName") or nested_agent.get("displayName") or "(no displayName)"
    state = view.get("state") or nested_agent.get("state") or "(unknown)"
    can_edit = permissions.get("canEdit")
    can_delete = permissions.get("canDelete")
    print(f"  - {name} | {display} | state={state} | canEdit={can_edit} canDelete={can_delete}")
    if name != "(unnamed)":
        print(f"      resource: {widget_skill_resource_name(cfg, name)}")
        print(f"      mention:  {widget_skill_mention(str(display), name)}")


def create_shell(
    s: requests.Session,
    cfg: LiveConfig,
    agent_id: str,
    instruction: str,
    display_name: str = DISPLAY_NAME,
    description: str = DESCRIPTION,
) -> dict:
    """Upsert the shell agent: create it, or if it already exists PATCH it in place.

    We deliberately do NOT delete-then-create. A skill agent delete is a *soft* delete that reserves
    the agent id (GET/PATCH return 404 while create returns 409 ALREADY_EXISTS) until the tombstone
    purges — so delete+create can strand the stable id our skillsSpec/mention marker depends on.
    Updating in place keeps the id (and never tombstones it). See --replace for the rare hard reset.
    """
    body = {
        "displayName": display_name,
        "description": description,
        "skillAgentDefinition": {"instruction": instruction},
    }
    r = s.post(f"{API}/{cfg.assistant}/agents", params={"agentId": agent_id}, json=body, timeout=HTTP_TIMEOUT)
    if r.status_code == 409:
        print("     agent exists — updating in place (no delete, keeps the id)")
        pr = s.patch(
            f"{API}/{cfg.assistant}/agents/{agent_id}",
            params={"updateMask": "displayName,description,skillAgentDefinition.instruction"},
            json=body,
            timeout=HTTP_TIMEOUT,
        )
        if pr.status_code == 404:
            # create=409 but patch=404 → the id is soft-deleted (tombstoned) and reserved. It cannot
            # be created or updated until the tombstone purges. Recover by pointing at a fresh id.
            raise SystemExit(
                f"Agent id '{agent_id}' is tombstoned (soft-deleted but reserved): create returns 409 "
                "while get/patch return 404. It was almost certainly hit by a prior delete/--replace.\n"
                "Recover by either: (a) wait for the tombstone to purge, then re-run; or (b) give the "
                f"skill a fresh agent id (e.g. the slug 'm365-command-planner') in packages/web-shell/.env "
                "(VITE_GE_COMMAND_PLANNER_SKILL / VITE_GE_SURFACE_COMMANDER_SKILL) and re-run bootstrap."
            )
        _raise_for_status_with_body(pr, "agents.patch")
        return pr.json()
    _raise_for_status_with_body(r, "agents.create")
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
    _raise_for_status_with_body(r, "agents.files:upload")
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
    _raise_for_status_with_body(r, "agents.patch")
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
    resource = widget_skill_resource_name(cfg, agent.get("name", agent_name))
    print(
        "    "
        + f'"skillsSpec": {{"skills": [{{"name": "{resource}"}}]}}'
    )
    print("  mention marker:")
    print("    " + widget_skill_mention(str(agent.get("displayName") or agent_name), agent.get("name", agent_name)))


def show(s: requests.Session, cfg: LiveConfig, agent_id: str) -> None:
    r = s.get(f"{API}/{cfg.assistant}/agents/{agent_id}", timeout=HTTP_TIMEOUT)
    _raise_for_status_with_body(r, "agents.get")
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
        help=(
            "widget matches the Gemini Enterprise web UI; public uses documented OAuth API "
            "(legacy is a backwards-compatible alias for public)"
        ),
    )
    ap.add_argument(
        "--list",
        action="store_true",
        help="list agents/skills and exit; in widget mode this lists visible skill agent views",
    )
    ap.add_argument(
        "--widget-origin",
        choices=[*WIDGET_AGENT_ORIGINS, "ALL"],
        default="USER",
        help="widget --list origin: USER private skills, GOOGLE built-ins, or ALL",
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
        raise SystemExit("--replace is only implemented for --api-mode public")
    if args.api_mode == "widget" and args.share:
        raise SystemExit("--share is only implemented for --api-mode public")
    if args.upload_existing and args.api_mode != "widget":
        raise SystemExit("--upload-existing is only valid with --api-mode widget")
    if args.upload_existing and args.single_file:
        raise SystemExit("--upload-existing only applies to zip bundle upload")

    # Resolve the live target up front so even a dry-run shows the real (env-provided) plan and a
    # live run cannot pick up accidental defaults.
    cfg = resolve_live_config()
    widget = resolve_widget_config() if args.api_mode == "widget" else None

    method = "A (single-file)" if args.single_file else "B (bundle upload)"
    op = "list agents" if args.list else f"provision agent '{args.agent_id}' via Method {method}"
    print(f"Plan: {op} ({_api_mode_label(args.api_mode)} API)")
    _print_target(cfg, args.agent_id, "PROVISION")
    if widget:
        print(f"      widget_config:  {widget.config_id}")
        if args.upload_existing:
            print("  step: --upload-existing -> skip create and upload to this numeric agent name")
    if args.replace:
        print("  step: --replace -> delete existing agent first")
    if args.share:
        print("  step: --share -> sharingConfig.scope=ALL_USERS (visible to ALL users in tenant)")
    if args.list:
        if args.api_mode == "widget":
            print("  step: --list -> widgetListAvailableAgentViews; no delete/create/upload")
        else:
            print("  step: --list -> GET agents under this assistant; no delete/create/upload")

    if not args.live:
        print("\nDRY-RUN (default): no API calls made. Re-run with --live to execute.")
        return 0

    s = session(cfg, args.api_mode)

    if args.list and args.api_mode == "widget":
        assert widget is not None
        origins = WIDGET_AGENT_ORIGINS if args.widget_origin == "ALL" else (args.widget_origin,)
        total = 0
        for origin in origins:
            views = list_widget_agent_views(s, cfg, widget, agent_origin=origin)
            total += len(views)
            print(f"  widget {origin} skill agent views: {len(views)}")
            for view in views:
                _print_widget_agent_view(cfg, view)
        print(f"  total: {total}")
        return 0

    if args.list:
        d = list_agents(s, cfg)
        agents = d.get("agents", [])
        if not isinstance(agents, list):
            print(json.dumps(d, indent=2, sort_keys=True))
            return 0
        print(f"  agents: {len(agents)}")
        for agent in agents:
            if not isinstance(agent, dict):
                continue
            print(
                "  - "
                + str(agent.get("name", "(unnamed)"))
                + f" | {agent.get('displayName', '(no displayName)')}"
                + f" | state={agent.get('state', '(unknown)')}"
            )
        return 0

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
