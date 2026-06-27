#!/usr/bin/env python3
"""
Create a Gemini Enterprise skill programmatically, using the **authenticated** equivalent of the
GE web UI's import flow (which we triangulated from captured browser traffic + a live end-to-end
test).

The `agents` resource is NOT in the public discovery doc, but the authenticated REST endpoints on
discoveryengine.googleapis.com work (GET/POST/DELETE all verified). Two methods:

  Method A — single-file skill (instruction only):
      POST {assistant}/agents?agentId=<id>
        { displayName, description, skillAgentDefinition: { instruction: "<full markdown>" } }

  Method B — multi-file bundle (SKILL.md + references/ + scripts/ + assets/):  [default]
      1) POST {assistant}/agents?agentId=<id>   with a placeholder skillAgentDefinition.instruction
      2) POST /upload/v1alpha/{assistant}/agents/<id>/files:upload?upload_protocol=raw
              Content-Type: application/zip, body = the zip bytes
         -> the server unpacks the zip: SKILL.md body -> instruction, the rest -> subfiles.
      3) GET to verify.

This is the same result the UI produces (create -> files:upload -> getAgentView), with a plain
OAuth Bearer token (ADC) instead of SAPISIDHASH/widget config.

Auth: ADC (gcloud auth print-access-token). Needs agents create/update on the engine.

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
  python3 create_skill.py --live                       # Method B: zip in ./m365-surface-commander.zip
  python3 create_skill.py --live --zip path/skill.zip  # Method B with a specific zip
  python3 create_skill.py --live --single-file SKILL.md  # Method A: inline instruction from markdown
  python3 create_skill.py --live --replace --yes       # delete an existing agent of the same id first
  python3 create_skill.py --live --share --yes         # set sharingConfig.scope=ALL_USERS after create
"""

import argparse
import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

import requests
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

# Request hardening: every call gets a finite timeout and a small, bounded retry budget.
HTTP_TIMEOUT = 60  # seconds
MAX_RETRIES = 3

REQUIRED_ENV = ("GE_PROJECT", "GE_PROJECT_NUMBER", "GE_ENGINE")


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


def session(cfg: LiveConfig) -> requests.Session:
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


def delete_agent(s: requests.Session, cfg: LiveConfig, agent_id: str):
    r = s.delete(f"{API}/{cfg.assistant}/agents/{agent_id}", timeout=HTTP_TIMEOUT)
    print(f"  delete {agent_id}: HTTP {r.status_code}")
    # 404 is fine (nothing to delete, idempotent); surface every other failure.
    if r.status_code == 404:
        return r
    r.raise_for_status()
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
    r.raise_for_status()
    return r.json()


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


def share(s: requests.Session, cfg: LiveConfig, agent_id: str):
    r = s.patch(
        f"{API}/{cfg.assistant}/agents/{agent_id}",
        params={"updateMask": "sharingConfig"},
        json={"sharingConfig": {"scope": "ALL_USERS"}},
        timeout=HTTP_TIMEOUT,
    )
    print(f"  share {agent_id}: HTTP {r.status_code}")
    r.raise_for_status()
    return r


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

    # Resolve the live target up front so even a dry-run shows the real (env-provided) plan and a
    # live run cannot pick up accidental defaults.
    cfg = resolve_live_config()

    method = "A (single-file)" if args.single_file else "B (bundle upload)"
    print(f"Plan: provision agent '{args.agent_id}' via Method {method}")
    _print_target(cfg, args.agent_id, "PROVISION")
    if args.replace:
        print("  step: --replace -> delete existing agent first")
    if args.share:
        print("  step: --share -> sharingConfig.scope=ALL_USERS (visible to ALL users in tenant)")

    if not args.live:
        print("\nDRY-RUN (default): no API calls made. Re-run with --live to execute.")
        return 0

    s = session(cfg)
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
