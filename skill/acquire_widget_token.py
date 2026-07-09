#!/usr/bin/env python3
"""
Probe the documented WidgetService.WidgetAcquireAccessToken RPC without leaking returned tokens.

Important: the official RPC says this method is deprecated/internal and proxies
DataConnectorService.AcquireAccessToken. It is therefore a probe for whether the response `uToken`
is a usable widget JWT; it is not assumed to be the durable replacement for copying the short-lived
widget Bearer from an authenticated Gemini Enterprise request.

The script prints only redacted diagnostics. If `--write-widget-token-file` is provided, it writes
only a returned `uToken` that validates as the Vertex AI Search widget JWT:

  iss = https://vertexaisearch.cloud.google
  aud = https://content-discoveryengine.googleapis.com
"""

from __future__ import annotations

import argparse
import json
import os
import re
import stat
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests

import create_skill


CONTENT_API = "https://content-discoveryengine.googleapis.com/v1alpha"
PUBLIC_API = "https://discoveryengine.googleapis.com/v1alpha"
DEFAULT_TOKEN_FILE = Path("/tmp/ge-widget-token")
TOKEN_RE = re.compile(
    r"\b(eyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_.=-]{10,}"
    r"|ya29\.[A-Za-z0-9_.=-]+)\b"
)


@dataclass(frozen=True)
class AcquireProbeConfig:
    location: str
    config_id: str
    connector_name: str
    scope: str | None = None
    action: str | None = None
    server_token: str | None = None


def connector_name_from_id(project_number: str, location: str, connector_id: str) -> str:
    return f"projects/{project_number}/locations/{location}/collections/{connector_id}/dataConnector"


def widget_acquire_access_token_payload(cfg: AcquireProbeConfig) -> dict[str, Any]:
    acquire_request: dict[str, str] = {"name": cfg.connector_name}
    if cfg.scope:
        acquire_request["scope"] = cfg.scope
    if cfg.action:
        acquire_request["action"] = cfg.action
    return {
        "location": f"locations/{cfg.location}",
        "configId": cfg.config_id,
        "acquireAccessTokenRequest": acquire_request,
        "additionalParams": {"token": "-", "origin": "ORIGIN_UNSPECIFIED"},
    }


def _erased_widget_connector_name(connector_name: str) -> str:
    match = re.search(r"/collections/([^/]+/dataConnector)$", connector_name)
    return f"collections/{match.group(1)}" if match else connector_name


def _acquire_request(cfg: AcquireProbeConfig, connector_name: str | None = None) -> dict[str, str]:
    acquire_request: dict[str, str] = {"name": connector_name or cfg.connector_name}
    if cfg.scope:
        acquire_request["scope"] = cfg.scope
    if cfg.action:
        acquire_request["action"] = cfg.action
    return acquire_request


def widget_acquire_access_token_variant_payloads(cfg: AcquireProbeConfig) -> list[tuple[str, dict[str, Any]]]:
    """Payload variants for live reverse-compat probing; all outputs are safe to print."""
    erased = _erased_widget_connector_name(cfg.connector_name)
    additional_params = {"token": "-", "origin": "ORIGIN_UNSPECIFIED"}

    def direct(location: str | None, connector: str) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "configId": cfg.config_id,
            "acquireAccessTokenRequest": _acquire_request(cfg, connector),
            "additionalParams": additional_params,
        }
        if location is not None:
            payload["location"] = location
        return payload

    def wrapped(location: str | None, connector: str) -> dict[str, Any]:
        request: dict[str, Any] = {
            "acquireAccessTokenRequest": _acquire_request(cfg, connector),
        }
        if location is not None:
            request["location"] = location
        return {
            "configId": cfg.config_id,
            "additionalParams": additional_params,
            "widgetAcquireAccessTokenRequest": request,
        }

    return [
        ("direct-resource-location-full-connector", direct(f"locations/{cfg.location}", cfg.connector_name)),
        ("direct-raw-location-full-connector", direct(cfg.location, cfg.connector_name)),
        ("direct-no-location-full-connector", direct(None, cfg.connector_name)),
        ("direct-resource-location-erased-connector", direct(f"locations/{cfg.location}", erased)),
        ("direct-raw-location-erased-connector", direct(cfg.location, erased)),
        ("wrapper-resource-location-full-connector", wrapped(f"locations/{cfg.location}", cfg.connector_name)),
        ("wrapper-raw-location-full-connector", wrapped(cfg.location, cfg.connector_name)),
        ("wrapper-no-location-full-connector", wrapped(None, cfg.connector_name)),
        ("wrapper-resource-location-erased-connector", wrapped(f"locations/{cfg.location}", erased)),
        ("wrapper-raw-location-erased-connector", wrapped(cfg.location, erased)),
    ]


def widget_acquire_access_token_url(location: str, host: str) -> str:
    base = CONTENT_API if host == "content" else PUBLIC_API
    return f"{base}/locations/{location}/widgetAcquireAccessToken"


def widget_acquire_access_token(
    s: requests.Session,
    cfg: AcquireProbeConfig,
    *,
    google_oauth_token: str,
    host: str = "content",
    quota_project: str | None = None,
) -> requests.Response:
    headers = {
        "Authorization": f"Bearer {google_oauth_token}",
        "Content-Type": "application/json",
        "Origin": create_skill.WIDGET_ORIGIN,
        "Referer": f"{create_skill.WIDGET_ORIGIN}/",
    }
    if cfg.server_token:
        headers["x-server-token"] = cfg.server_token
    if quota_project:
        headers["x-goog-user-project"] = quota_project
    return s.post(
        widget_acquire_access_token_url(cfg.location, host),
        json=widget_acquire_access_token_payload(cfg),
        headers=headers,
        timeout=create_skill.HTTP_TIMEOUT,
    )


def widget_acquire_access_token_with_payload(
    s: requests.Session,
    cfg: AcquireProbeConfig,
    *,
    google_oauth_token: str,
    payload: dict[str, Any],
    host: str = "content",
    quota_project: str | None = None,
) -> requests.Response:
    headers = {
        "Authorization": f"Bearer {google_oauth_token}",
        "Content-Type": "application/json",
        "Origin": create_skill.WIDGET_ORIGIN,
        "Referer": f"{create_skill.WIDGET_ORIGIN}/",
    }
    if cfg.server_token:
        headers["x-server-token"] = cfg.server_token
    if quota_project:
        headers["x-goog-user-project"] = quota_project
    return s.post(
        widget_acquire_access_token_url(cfg.location, host),
        json=payload,
        headers=headers,
        timeout=create_skill.HTTP_TIMEOUT,
    )


def token_payload(token: str) -> dict[str, Any] | None:
    try:
        return create_skill._jwt_payload(token)
    except Exception:
        return None


def token_summary(token: str | None) -> dict[str, Any]:
    if not token:
        return {"present": False}
    payload = token_payload(token)
    summary: dict[str, Any] = {"present": True, "length": len(token)}
    if payload is None:
        summary["kind"] = "opaque"
        return summary
    exp = payload.get("exp")
    expires_at = (
        datetime.fromtimestamp(float(exp), tz=timezone.utc).isoformat()
        if isinstance(exp, (int, float))
        else None
    )
    summary.update(
        {
            "kind": "jwt",
            "iss": payload.get("iss"),
            "aud": payload.get("aud"),
            "sub": payload.get("sub"),
            "exp": exp,
            "expiresAt": expires_at,
            "secondsRemaining": int(float(exp) - time.time())
            if isinstance(exp, (int, float))
            else None,
            "isWidgetBearer": is_widget_bearer(token),
        }
    )
    return summary


def is_widget_bearer(token: str | None) -> bool:
    if not token:
        return False
    try:
        create_skill._validate_widget_bearer_token(token)
        return True
    except SystemExit:
        return False


def redact_secrets(text: str) -> str:
    return TOKEN_RE.sub("<redacted-token>", text)


def response_diagnostics(response: requests.Response) -> dict[str, Any]:
    out: dict[str, Any] = {"status": response.status_code}
    try:
        body = response.json()
    except Exception:
        out["body"] = redact_secrets(response.text[:2000])
        return out
    if not isinstance(body, dict):
        out["bodyType"] = type(body).__name__
        return out
    access = body.get("acquireAccessTokenResponse", {})
    if not isinstance(access, dict):
        access = {}
    out["keys"] = sorted(body.keys())
    out["uToken"] = token_summary(body.get("uToken") if isinstance(body.get("uToken"), str) else None)
    out["connectorAccessToken"] = token_summary(
        access.get("accessToken") if isinstance(access.get("accessToken"), str) else None
    )
    if response.status_code >= 400:
        out["error"] = redact_secrets(json.dumps(body, indent=2, sort_keys=True)[:2000])
    return out


def write_token(path: Path, token: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(token, encoding="utf-8")
    path.chmod(stat.S_IRUSR | stat.S_IWUSR)


def google_oauth_token(args: argparse.Namespace) -> str:
    token = (args.google_token or os.environ.get("GE_GOOGLE_OAUTH_TOKEN") or "").strip()
    token_file = args.google_token_file or os.environ.get("GE_GOOGLE_OAUTH_TOKEN_FILE")
    if not token and token_file:
        token = Path(token_file).read_text(encoding="utf-8").strip()
    if token:
        return token.removeprefix("Bearer ").strip()
    if not args.use_gcloud:
        raise SystemExit(
            "Missing Google OAuth token. Pass --google-token, --google-token-file, "
            "or --use-gcloud."
        )
    env = os.environ.copy()
    if args.cloudsdk_config:
        env["CLOUDSDK_CONFIG"] = args.cloudsdk_config
    proc = subprocess.run(
        ["gcloud", "auth", "print-access-token"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env=env,
    )
    if proc.returncode != 0:
        raise SystemExit(
            "gcloud auth print-access-token failed; re-login is required.\n"
            + redact_secrets(proc.stderr.strip())
        )
    token = proc.stdout.strip()
    if not token:
        raise SystemExit("gcloud returned an empty access token.")
    return token


def quota_project(args: argparse.Namespace) -> str | None:
    return (
        args.quota_project
        or os.environ.get("GE_QUOTA_PROJECT")
        or os.environ.get("GOOGLE_CLOUD_QUOTA_PROJECT")
        or os.environ.get("GOOGLE_CLOUD_PROJECT")
        or os.environ.get("GE_PROJECT")
    )


def resolve_probe_config(args: argparse.Namespace) -> AcquireProbeConfig:
    location = args.location or os.environ.get("GE_LOCATION") or create_skill.LOCATION
    config_id = args.config_id or os.environ.get("GE_WIDGET_CONFIG_ID")
    if not config_id:
        raise SystemExit("Set GE_WIDGET_CONFIG_ID or pass --config-id.")
    connector_name = args.connector_name or os.environ.get("GE_WIDGET_CONNECTOR_NAME")
    connector_id = args.connector_id or os.environ.get("GE_WIDGET_CONNECTOR_ID")
    if not connector_name and connector_id:
        project_number = args.project_number or os.environ.get("GE_PROJECT_NUMBER")
        if not project_number:
            raise SystemExit("Set GE_PROJECT_NUMBER or pass --project-number with --connector-id.")
        connector_name = connector_name_from_id(project_number, location, connector_id)
    if not connector_name:
        raise SystemExit(
            "Set GE_WIDGET_CONNECTOR_NAME or pass --connector-name. "
            "Example: projects/<project-number>/locations/global/collections/"
            "msft-onedrive-fed_1779469629030/dataConnector"
        )
    return AcquireProbeConfig(
        location=location,
        config_id=config_id,
        connector_name=connector_name,
        scope=args.scope,
        action=args.action,
        server_token=args.server_token or os.environ.get("GE_WIDGET_SERVER_TOKEN"),
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Probe WidgetService.WidgetAcquireAccessToken and report whether response.uToken is a "
            "usable content-discoveryengine widget JWT. Tokens are never printed."
        )
    )
    parser.add_argument("--live", action="store_true", help="perform the HTTP request")
    parser.add_argument("--host", choices=("content", "public", "both"), default="content")
    parser.add_argument("--config-id")
    parser.add_argument("--location")
    parser.add_argument("--connector-name")
    parser.add_argument("--connector-id")
    parser.add_argument("--project-number")
    parser.add_argument("--scope")
    parser.add_argument("--action")
    parser.add_argument("--server-token")
    parser.add_argument("--google-token")
    parser.add_argument("--google-token-file")
    parser.add_argument("--use-gcloud", action="store_true")
    parser.add_argument("--quota-project")
    parser.add_argument(
        "--probe-variants",
        action="store_true",
        help="try documented/direct/widget-wrapper payload variants without printing token values",
    )
    parser.add_argument(
        "--cloudsdk-config",
        default=os.environ.get("CLOUDSDK_CONFIG", "/home/user/ge-msft/.gcloud"),
    )
    parser.add_argument("--write-widget-token-file", type=Path, default=None)
    args = parser.parse_args(argv)

    cfg = resolve_probe_config(args)
    payload = widget_acquire_access_token_payload(cfg)
    print("WidgetAcquireAccessToken probe")
    print(f"  host:       {args.host}")
    print(f"  location:   {cfg.location}")
    print(f"  config_id:  {cfg.config_id}")
    print(f"  connector:  {cfg.connector_name}")
    print("  payload:    " + json.dumps(payload, sort_keys=True))

    if not args.live:
        print("\nDRY-RUN: no network call made. Re-run with --live to execute.")
        return 0

    oauth_token = google_oauth_token(args)
    user_project = quota_project(args)
    hosts = ("content", "public") if args.host == "both" else (args.host,)
    exit_code = 1
    variants = (
        widget_acquire_access_token_variant_payloads(cfg)
        if args.probe_variants
        else [("documented", payload)]
    )
    for host in hosts:
        s = requests.Session()
        for variant_name, variant_payload in variants:
            response = widget_acquire_access_token_with_payload(
                s,
                cfg,
                google_oauth_token=oauth_token,
                payload=variant_payload,
                host=host,
                quota_project=user_project,
            )
            diagnostics = response_diagnostics(response)
            print(f"\n{host} endpoint diagnostics ({variant_name}):")
            print(json.dumps(diagnostics, indent=2, sort_keys=True))
            if response.status_code < 400:
                body = response.json()
                u_token = body.get("uToken") if isinstance(body, dict) else None
                if isinstance(u_token, str) and is_widget_bearer(u_token):
                    exit_code = 0
                    if args.write_widget_token_file:
                        write_token(args.write_widget_token_file, u_token)
                        print(
                            f"wrote validated widget bearer token to {args.write_widget_token_file}"
                        )
                else:
                    print("response did not contain a validated widget bearer uToken")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
