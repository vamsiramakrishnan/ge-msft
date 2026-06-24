#!/usr/bin/env python3
"""
Create a Gemini Enterprise skill programmatically, using the **authenticated** equivalent of the
GE web UI's import flow (which we triangulated from captured browser traffic + a live end-to-end
test against phoenix-telco).

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

Usage:
  python3 create_skill.py                          # Method B: zip in ./m365-surface-commander.zip
  python3 create_skill.py --zip path/to/skill.zip  # Method B with a specific zip
  python3 create_skill.py --single-file SKILL.md   # Method A: inline instruction from a markdown file
  python3 create_skill.py --replace                # delete an existing agent of the same id first
  python3 create_skill.py --share                  # set sharingConfig.scope=ALL_USERS after create
"""

import argparse
import os
import sys
from pathlib import Path

import google.auth
import google.auth.transport.requests
import requests

# Configure for your engine via env vars (defaults shown). For upstream use, set GE_PROJECT etc.
PROJECT = os.environ.get("GE_PROJECT", "vital-octagon-19612")
PROJECT_NUMBER = os.environ.get("GE_PROJECT_NUMBER", "440790012685")
LOCATION = os.environ.get("GE_LOCATION", "global")
ENGINE = os.environ.get("GE_ENGINE", "phoenix-telco_1751440131886")
AGENT_ID = os.environ.get("GE_AGENT_ID", "m365-surface-commander")  # slug referenced in skillsSpec
DISPLAY_NAME = "M365 Surface Commander"
DESCRIPTION = (
    "Reads, analyzes, and edits the Microsoft 365 document the user has open via a "
    "compact command-line protocol the Office add-in applies as reviewable changes."
)

ASSISTANT = (
    f"projects/{PROJECT_NUMBER}/locations/{LOCATION}"
    f"/collections/default_collection/engines/{ENGINE}/assistants/default_assistant"
)
API = "https://discoveryengine.googleapis.com/v1alpha"
UPLOAD_API = "https://discoveryengine.googleapis.com/upload/v1alpha"


def session() -> requests.Session:
    creds, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
    creds.refresh(google.auth.transport.requests.Request())
    s = requests.Session()
    s.headers.update({
        "Authorization": f"Bearer {creds.token}",
        "X-Goog-User-Project": PROJECT,
    })
    return s


def delete_if_exists(s: requests.Session, agent_id: str) -> None:
    r = s.delete(f"{API}/{ASSISTANT}/agents/{agent_id}")
    print(f"  delete {agent_id}: HTTP {r.status_code}")


def create_shell(s: requests.Session, agent_id: str, instruction: str) -> dict:
    r = s.post(
        f"{API}/{ASSISTANT}/agents",
        params={"agentId": agent_id},
        json={
            "displayName": DISPLAY_NAME,
            "description": DESCRIPTION,
            "skillAgentDefinition": {"instruction": instruction},
        },
    )
    r.raise_for_status()
    return r.json()


def upload_zip(s: requests.Session, agent_id: str, zip_path: Path) -> dict:
    r = s.post(
        f"{UPLOAD_API}/{ASSISTANT}/agents/{agent_id}/files:upload",
        params={"upload_protocol": "raw"},
        headers={"Content-Type": "application/zip"},
        data=zip_path.read_bytes(),
    )
    r.raise_for_status()
    return r.json()


def share(s: requests.Session, agent_id: str) -> None:
    r = s.patch(
        f"{API}/{ASSISTANT}/agents/{agent_id}",
        params={"updateMask": "sharingConfig"},
        json={"sharingConfig": {"scope": "ALL_USERS"}},
    )
    print(f"  share {agent_id}: HTTP {r.status_code}")


def show(s: requests.Session, agent_id: str) -> None:
    d = s.get(f"{API}/{ASSISTANT}/agents/{agent_id}").json()
    sd = d.get("skillAgentDefinition", {})
    print(f"  name:        {d.get('name','').split('/')[-1]}")
    print(f"  displayName: {d.get('displayName')}")
    print(f"  state:       {d.get('state')}")
    print(f"  instruction: {len(sd.get('instruction',''))} chars")
    print(f"  subfiles:    {[f['fileName'] for f in sd.get('subfiles',[])] or '(none)'}")
    print(f"\n  reference in streamAssist via:")
    print(f'    "skillsSpec": {{"skills": [{{"name": "{d.get("name")}"}}]}}')


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--zip", default=str(Path(__file__).parent / "m365-surface-commander.zip"))
    ap.add_argument("--single-file", metavar="SKILL.md",
                    help="Method A: create with this file's contents as the inline instruction")
    ap.add_argument("--agent-id", default=AGENT_ID)
    ap.add_argument("--replace", action="store_true", help="delete an existing agent of this id first")
    ap.add_argument("--share", action="store_true", help="set sharingConfig.scope=ALL_USERS")
    args = ap.parse_args()

    s = session()
    if args.replace:
        delete_if_exists(s, args.agent_id)

    if args.single_file:
        print(f"Method A — single-file create from {args.single_file}")
        instruction = Path(args.single_file).read_text(encoding="utf-8")
        create_shell(s, args.agent_id, instruction)
    else:
        zip_path = Path(args.zip)
        if not zip_path.exists():
            sys.exit(f"zip not found: {zip_path}")
        print(f"Method B — bundle upload from {zip_path}")
        print("  1) create shell agent")
        create_shell(s, args.agent_id, "placeholder — replaced by SKILL.md on upload")
        print("  2) upload zip (server unpacks)")
        upload_zip(s, args.agent_id, zip_path)

    if args.share:
        share(s, args.agent_id)
    print("  3) verify")
    show(s, args.agent_id)


if __name__ == "__main__":
    main()
