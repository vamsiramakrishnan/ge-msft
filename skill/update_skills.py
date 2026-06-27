#!/usr/bin/env python3
"""
Batch delete/update the bundled Gemini Enterprise skills.

This is a thin, explicit wrapper around create_skill.py. It targets the two committed zip bundles:

  - m365-surface-commander.zip
  - m365-command-planner.zip

Dry-run is the default and performs no network calls. Live destructive operations require --yes.

Examples:
  python3 skill/update_skills.py
  GE_SURFACE_COMMANDER_AGENT_ID=740... GE_COMMAND_PLANNER_AGENT_ID=175... \
    python3 skill/update_skills.py --live --replace --yes
  GE_SURFACE_COMMANDER_AGENT_ID=740... GE_COMMAND_PLANNER_AGENT_ID=175... \
    python3 skill/update_skills.py --live --upload-existing
  python3 skill/update_skills.py --live --create-new
  python3 skill/update_skills.py --live --api-mode legacy --replace --yes
  python3 skill/update_skills.py --live --api-mode legacy --delete-only --yes
  python3 skill/update_skills.py --live --replace --yes --only m365-surface-commander
"""

from __future__ import annotations

import argparse
import hashlib
import os
import sys
from dataclasses import dataclass
from pathlib import Path

import create_skill

HERE = Path(__file__).parent


@dataclass(frozen=True)
class SkillBundle:
    agent_id: str
    zip_path: Path
    display_name: str
    description: str
    env_key: str
    label: str


BUNDLES = (
    SkillBundle(
        agent_id=os.environ.get("GE_SURFACE_COMMANDER_AGENT_ID", "m365-surface-commander"),
        zip_path=HERE / "m365-surface-commander.zip",
        display_name="M365 Surface Commander",
        description=(
            "Reads, analyzes, and edits the Microsoft 365 document the user has open via a "
            "compact command-line protocol the Office add-in applies as reviewable changes."
        ),
        env_key="VITE_GE_SURFACE_COMMANDER_SKILL",
        label="m365-surface-commander",
    ),
    SkillBundle(
        agent_id=os.environ.get("GE_COMMAND_PLANNER_AGENT_ID", "m365-command-planner"),
        zip_path=HERE / "m365-command-planner.zip",
        display_name="M365 Command Planner",
        description=(
            "Turns a user's free-text Microsoft 365 request into a structured, reviewable plan "
            "before the surface commander executes document commands."
        ),
        env_key="VITE_GE_COMMAND_PLANNER_SKILL",
        label="m365-command-planner",
    ),
)


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _select(names: list[str] | None) -> list[SkillBundle]:
    if not names:
        return list(BUNDLES)
    wanted = set(names)
    bundles = [b for b in BUNDLES if b.agent_id in wanted]
    missing = wanted - {b.agent_id for b in bundles}
    if missing:
        raise SystemExit(f"unknown --only skill(s): {', '.join(sorted(missing))}")
    return bundles


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Delete/update the bundled GE skill zip files.")
    ap.add_argument(
        "--only",
        action="append",
        choices=[b.agent_id for b in BUNDLES],
        help="limit to one bundled skill; repeatable",
    )
    ap.add_argument(
        "--live",
        action="store_true",
        help="actually call Discovery Engine (default is dry-run)",
    )
    ap.add_argument(
        "--api-mode",
        choices=create_skill.API_MODES,
        default="widget",
        help="widget matches the Gemini Enterprise web UI; legacy uses the older public API path",
    )
    ap.add_argument(
        "--upload-existing",
        action="store_true",
        help="widget mode: upload to the selected numeric agent IDs instead of creating new agents",
    )
    ap.add_argument(
        "--create-new",
        action="store_true",
        help="widget mode: create new numeric agents, upload the zips, and print the returned names",
    )
    ap.add_argument(
        "--replace",
        action="store_true",
        help="delete each existing agent first, then recreate and upload its zip (requires --yes)",
    )
    ap.add_argument(
        "--delete-only",
        action="store_true",
        help="delete each selected agent and stop; no upload (requires --yes)",
    )
    ap.add_argument(
        "--share",
        action="store_true",
        help="set sharingConfig.scope=ALL_USERS after upload (requires --yes)",
    )
    ap.add_argument(
        "--yes",
        action="store_true",
        help="confirm destructive or tenant-wide actions",
    )
    args = ap.parse_args(argv)

    destructive = []
    if args.replace:
        destructive.append("--replace")
    if args.delete_only:
        destructive.append("--delete-only")
    if args.share:
        destructive.append("--share")
    if destructive and not args.yes:
        raise SystemExit(
            f"Refusing {' and '.join(destructive)} without confirmation. "
            "Re-run with --yes once the dry-run target is correct."
        )
    if args.delete_only and args.share:
        raise SystemExit("--delete-only cannot be combined with --share")
    if args.upload_existing and args.create_new:
        raise SystemExit("--upload-existing and --create-new are mutually exclusive")
    if args.api_mode == "widget" and (args.delete_only or args.share):
        raise SystemExit("--delete-only and --share are only implemented for legacy mode")
    if args.api_mode == "legacy" and (args.upload_existing or args.create_new):
        raise SystemExit("--upload-existing/--create-new are widget mode options")

    cfg = create_skill.resolve_live_config()
    widget = create_skill.resolve_widget_config() if args.api_mode == "widget" else None
    bundles = _select(args.only)

    print(
        f"Plan: {'delete' if args.delete_only else 'update'} {len(bundles)} "
        f"Gemini Enterprise skill(s) via {args.api_mode} API"
    )
    create_skill._print_target(cfg, ",".join(b.agent_id for b in bundles), "BATCH")
    if widget:
        print(f"      widget_config:  {widget.config_id}")
    updated: list[tuple[SkillBundle, str]] = []
    for bundle in bundles:
        if not bundle.zip_path.exists() and not args.delete_only:
            raise SystemExit(f"zip not found: {bundle.zip_path}")
        print(f"\n  skill:       {bundle.agent_id}")
        print(f"  displayName: {bundle.display_name}")
        if not args.delete_only:
            print(f"  zip:         {bundle.zip_path}")
            print(f"  sha256:      {_sha256(bundle.zip_path)}")
        if args.replace:
            print("  step:        delete existing agent first")
        if args.delete_only:
            print("  step:        delete only; no upload")
        if args.share:
            print("  step:        share with ALL_USERS after upload")
        if args.api_mode == "widget":
            if args.replace:
                print("  step:        delete this numeric widget agent, create replacement, upload zip")
            elif args.create_new:
                print("  step:        create a new widget agent and upload zip")
            elif args.upload_existing:
                print("  step:        upload zip to this existing numeric widget agent")
            else:
                print(
                    "  step:        dry-run only; live widget mode needs --replace, "
                    "--upload-existing, or --create-new"
                )

    if not args.live:
        print("\nDRY-RUN (default): no API calls made. Re-run with --live to execute.")
        return 0

    if args.api_mode == "widget" and not (args.upload_existing or args.create_new or args.replace):
        raise SystemExit(
            "Widget live mode requires --replace (with numeric GE_*_AGENT_ID values), "
            "--upload-existing, or --create-new."
        )

    s = create_skill.session(cfg, args.api_mode)
    for bundle in bundles:
        if args.api_mode == "widget":
            assert widget is not None
            print(f"\nUpdate {bundle.agent_id}")
            if args.replace:
                create_skill._print_target(cfg, bundle.agent_id, "WIDGET DELETE")
                create_skill.delete_widget_agent(s, cfg, widget, bundle.agent_id)
                print("  1) create replacement widget shell agent")
                agent = create_skill.create_widget_agent(
                    s,
                    cfg,
                    widget,
                    "placeholder — replaced by SKILL.md on upload",
                    bundle.display_name,
                    bundle.description,
                )
                agent_name = agent["name"]
                print(f"     created agent.name {agent_name}")
            elif args.create_new:
                print("  1) create widget shell agent")
                agent = create_skill.create_widget_agent(
                    s,
                    cfg,
                    widget,
                    "placeholder — replaced by SKILL.md on upload",
                    bundle.display_name,
                    bundle.description,
                )
                agent_name = agent["name"]
                print(f"     created agent.name {agent_name}")
            else:
                agent_name = bundle.agent_id
                print(f"  1) use existing agent name {agent_name}")
            print("  2) resumable upload zip (server unpacks)")
            create_skill.upload_zip_resumable(s, cfg, agent_name, bundle.zip_path)
            print("  3) verify")
            create_skill.show_widget(s, cfg, widget, agent_name)
            updated.append((bundle, agent_name))
            continue

        if args.replace or args.delete_only:
            create_skill._print_target(cfg, bundle.agent_id, "DELETE")
            create_skill.delete_agent(s, cfg, bundle.agent_id)
        if args.delete_only:
            continue

        print(f"\nUpdate {bundle.agent_id}")
        print("  1) create shell agent")
        try:
            create_skill.create_shell(
                s,
                cfg,
                bundle.agent_id,
                "placeholder — replaced by SKILL.md on upload",
                bundle.display_name,
                bundle.description,
            )
        except Exception as exc:
            status = getattr(getattr(exc, "response", None), "status_code", None)
            if status != 409:
                raise
            print("     existing agent shell detected (HTTP 409); uploading zip in place")
        print("  2) upload zip (server unpacks)")
        create_skill.upload_zip(s, cfg, bundle.agent_id, bundle.zip_path)
        if args.share:
            create_skill._print_target(cfg, bundle.agent_id, "SHARE (ALL_USERS)")
            create_skill.share(s, cfg, bundle.agent_id)
        print("  3) verify")
        create_skill.show(s, cfg, bundle.agent_id)
        updated.append((bundle, bundle.agent_id))
    if updated:
        print("\nUpdated skill references for the add-in env:")
        for bundle, agent_name in updated:
            print(f"  {bundle.env_key}={bundle.label}={cfg.assistant}/agents/{agent_name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
