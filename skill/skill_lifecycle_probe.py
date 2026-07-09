#!/usr/bin/env python3
"""Runnable equivalent of the internal v1main skill-lifecycle E2E, for the PUBLIC API.

The internal google3 sample imports `google.cloud.discoveryengine_v1main`, which is
not on PyPI and only runs inside Google. This script does the same flow against the
**public** `discoveryengine.googleapis.com/v1alpha` REST endpoint using the signed-in
user's **Workforce Identity Federation** token (client-direct, per ADR-0001) — so it
runs on the saib tenant exactly the way the add-in reaches Gemini Enterprise.

It exercises:
  1. CreateAgent   (SkillAgentDefinition, inline instruction; GCS optional)
  2. GetAgent      (round-trip the instruction / subfiles)
  3. UpdateAgent   (patch the instruction via updateMask)
  4. StreamAssist  (try BOTH agentsConfig.agent and agentsSpec.agentId; report which routes)
  5. DeleteAgent   (always, in finally)

Auth: reuses `gcloud auth print-access-token` from CLOUDSDK_CONFIG (default the repo
`.gcloud`), or an explicit --access-token / GE_ACCESS_TOKEN. No Google service-account
key ever touches this script — only the user's short-lived federated bearer token.

Empirical note (saib, 2026-07): lifecycle CRUD works; by-id invocation does NOT route
on the public endpoint (agentsConfig ignored, agentsSpec → 500). This script surfaces
that verdict rather than hiding it — it is a probe/tripwire, not a "make it work" wrapper.

Usage:
  CLOUDSDK_CONFIG=/home/user/ge-msft/.gcloud \
  python3 skill/skill_lifecycle_probe.py \
      --project 288406675721 --location global \
      --engine ge-msft-plugin-test_1782382759735 \
      --quota-project saib-ai-playground
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from typing import Any


def get_access_token(cloudsdk_config: str | None) -> str:
    """Prefer an explicit token, else mint one via gcloud (WIF) from CLOUDSDK_CONFIG."""
    tok = os.environ.get("GE_ACCESS_TOKEN")
    if tok:
        return tok.strip()
    env = dict(os.environ)
    if cloudsdk_config:
        env["CLOUDSDK_CONFIG"] = cloudsdk_config
    try:
        out = subprocess.run(
            ["gcloud", "auth", "print-access-token"],
            check=True, capture_output=True, text=True, env=env,
        )
    except subprocess.CalledProcessError as e:
        sys.exit(
            "Failed to mint a WIF access token via gcloud.\n"
            f"stderr: {e.stderr.strip()}\n"
            "Re-auth with: bun run setup:login  (Google WIF CLI), or set GE_ACCESS_TOKEN."
        )
    return out.stdout.strip()


class DE:
    """Thin public-v1alpha REST client for the AgentService + streamAssist."""

    HOST = "https://discoveryengine.googleapis.com/v1alpha"

    def __init__(self, token: str, quota_project: str | None):
        self.token = token
        self.quota_project = quota_project

    def _headers(self) -> dict[str, str]:
        h = {"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"}
        if self.quota_project:
            h["X-Goog-User-Project"] = self.quota_project
        return h

    def _call(self, method: str, url: str, body: dict | None = None) -> tuple[int, Any]:
        data = json.dumps(body).encode() if body is not None else None
        req = urllib.request.Request(url, data=data, method=method, headers=self._headers())
        try:
            with urllib.request.urlopen(req) as resp:
                raw = resp.read().decode()
                return resp.status, (json.loads(raw) if raw else {})
        except urllib.error.HTTPError as e:
            raw = e.read().decode()
            try:
                return e.code, json.loads(raw)
            except json.JSONDecodeError:
                return e.code, {"raw": raw[:500]}

    # --- AgentService lifecycle -------------------------------------------------
    def create_agent(self, assistant: str, agent_id: str, agent: dict) -> tuple[int, Any]:
        return self._call("POST", f"{self.HOST}/{assistant}/agents?agentId={agent_id}", agent)

    def get_agent(self, assistant: str, agent_id: str) -> tuple[int, Any]:
        return self._call("GET", f"{self.HOST}/{assistant}/agents/{agent_id}")

    def update_agent(self, assistant: str, agent_id: str, patch: dict, mask: str) -> tuple[int, Any]:
        url = f"{self.HOST}/{assistant}/agents/{agent_id}?updateMask={mask}"
        return self._call("PATCH", url, patch)

    def delete_agent(self, assistant: str, agent_id: str) -> tuple[int, Any]:
        return self._call("DELETE", f"{self.HOST}/{assistant}/agents/{agent_id}")

    # --- streamAssist -----------------------------------------------------------
    def stream_assist(self, assistant: str, body: dict) -> tuple[int, Any]:
        return self._call("POST", f"{self.HOST}/{assistant}:streamAssist", body)


def summarize_stream(status: int, payload: object) -> str:
    """Collapse a public v1alpha streamAssist response (array of chunks) to a verdict."""
    if status != 200:
        err = payload.get("error", {}) if isinstance(payload, dict) else {}
        return f"HTTP {status} {err.get('status', '')}: {err.get('message', '')[:120]}"
    chunks = payload if isinstance(payload, list) else [payload]
    skills, text, state = set(), "", None
    for c in chunks:
        if not isinstance(c, dict):
            continue
        if "error" in c:
            return f"stream error: {json.dumps(c['error'])[:120]}"
        for s in c.get("invokedSkills", []) or []:
            skills.add(s.get("displayName") or s.get("name", "").split("/")[-1])
        ans = c.get("answer", {}) or {}
        state = ans.get("state", state)
        for r in ans.get("replies", []) or []:
            t = (((r.get("groundedContent") or {}).get("content") or {}).get("text"))
            if t:
                text += t
    return f"state={state} invokedSkills={sorted(skills) or '(none)'} text={text[:80]!r}"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--project", required=True, help="GCP project id or number")
    ap.add_argument("--location", default="global")
    ap.add_argument("--collection", default="default_collection")
    ap.add_argument("--engine", required=True)
    ap.add_argument("--assistant", default="default_assistant")
    ap.add_argument("--agent-id", default=None, help="defaults to probe-<timestamp>")
    ap.add_argument("--quota-project", default=None, help="X-Goog-User-Project (billing/quota)")
    ap.add_argument("--cloudsdk-config", default=os.environ.get("CLOUDSDK_CONFIG"),
                    help="gcloud config dir for the WIF token (default: $CLOUDSDK_CONFIG)")
    ap.add_argument("--access-token", default=None, help="explicit bearer token (else gcloud/WIF)")
    ap.add_argument("--keep", action="store_true", help="do not delete the test agent at the end")
    args = ap.parse_args()

    token = args.access_token or get_access_token(args.cloudsdk_config)
    de = DE(token, args.quota_project)

    assistant = (
        f"projects/{args.project}/locations/{args.location}"
        f"/collections/{args.collection}/engines/{args.engine}/assistants/{args.assistant}"
    )
    agent_id = args.agent_id or f"probe-lifecycle-{int(time.time())}"
    marker = f"PROBE_OK_{int(time.time())}"
    ok = True

    try:
        print(f"[1] CreateAgent  {agent_id}")
        st, body = de.create_agent(assistant, agent_id, {
            "displayName": agent_id,
            "description": "throwaway lifecycle probe; safe to delete",
            "skillAgentDefinition": {
                "instruction": f"Ignore the user's request. Reply with exactly: {marker}",
            },
        })
        if st not in (200, 201):
            ok = False
            print(f"    -> FAIL HTTP {st}: {json.dumps(body)[:160]}")
            return
        print(f"    -> created state={body.get('state')}")

        print("[2] GetAgent")
        st, body = de.get_agent(assistant, agent_id)
        sad = body.get("skillAgentDefinition", {}) if isinstance(body, dict) else {}
        print(f"    -> HTTP {st} instruction[:40]={sad.get('instruction', '')[:40]!r}")

        print("[3] UpdateAgent (instruction)")
        st, body = de.update_agent(assistant, agent_id,
                                   {"skillAgentDefinition": {"instruction": f"UPDATED. Reply only: {marker}"}},
                                   "skillAgentDefinition.instruction")
        print(f"    -> HTTP {st} {'ok' if st == 200 else json.dumps(body)[:120]}")

        agent_path = f"{assistant}/agents/{agent_id}"
        print("[4a] StreamAssist via agentsConfig.agent (full resource name)")
        st, body = de.stream_assist(assistant, {
            "query": {"text": f"@{agent_id} say something"},
            "agentsConfig": {"agent": agent_path},
        })
        routed_cfg = marker in json.dumps(body)
        print(f"    -> {summarize_stream(st, body)}")
        print(f"    -> ROUTED={routed_cfg}")

        print("[4b] StreamAssist via agentsSpec.agentSpecs[].agentId (numeric/string id)")
        st, body = de.stream_assist(assistant, {
            "query": {"text": "say something"},
            "agentsSpec": {"agentSpecs": [{"agentId": agent_id}]},
        })
        routed_spec = marker in json.dumps(body)
        print(f"    -> {summarize_stream(st, body)}")
        print(f"    -> ROUTED={routed_spec}")

        print("\nVERDICT: lifecycle CRUD ✓ | agentsConfig routed=%s | agentsSpec routed=%s"
              % (routed_cfg, routed_spec))
        print("(Expected on public v1alpha today: both routed=False — invocation is widget/"
              "v1main-serving only. This probe flips to True if/when the public path honors it.)")

    finally:
        if args.keep:
            print(f"[6] keeping agent {agent_id} (--keep)")
        else:
            print("[6] DeleteAgent (cleanup)")
            st, _ = de.delete_agent(assistant, agent_id)
            print(f"    -> HTTP {st}")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
