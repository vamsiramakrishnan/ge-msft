#!/usr/bin/env python3
"""
Batch delete/update the bundled Gemini Enterprise skills.

This is a thin, explicit wrapper around create_skill.py. It targets the two committed zip bundles:

  - m365-surface-commander.zip
  - m365-command-planner.zip

Dry-run is the default and performs no network calls. Live destructive operations require --yes.

Examples:
  python3 skill/update_skills.py
  python3 skill/update_skills.py --live --replace --yes
  python3 skill/update_skills.py --live --delete-only --yes
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


BUNDLES = (
    SkillBundle(
        agent_id=os.environ.get("GE_SURFACE_COMMANDER_AGENT_ID", "m365-surface-commander"),
        zip_path=HERE / "m365-surface-commander.zip",
        display_name="M365 Surface Commander",
        description=(
            "Reads, analyzes, and edits the Microsoft 365 document the user has open via a "
            "compact command-line protocol the Office add-in applies as reviewable changes."
        ),
    ),
    SkillBundle(
        agent_id=os.environ.get("GE_COMMAND_PLANNER_AGENT_ID", "m365-command-planner"),
        zip_path=HERE / "m365-command-planner.zip",
        display_name="M365 Command Planner",
        description=(
            "Turns a user's free-text Microsoft 365 request into a structured, reviewable plan "
            "before the surface commander executes document commands."
        ),
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

    cfg = create_skill.resolve_live_config()
    bundles = _select(args.only)

    print(f"Plan: {'delete' if args.delete_only else 'update'} {len(bundles)} Gemini Enterprise skill(s)")
    create_skill._print_target(cfg, ",".join(b.agent_id for b in bundles), "BATCH")
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

    if not args.live:
        print("\nDRY-RUN (default): no API calls made. Re-run with --live to execute.")
        return 0

    s = create_skill.session(cfg)
    for bundle in bundles:
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
    return 0


if __name__ == "__main__":
    sys.exit(main())
