#!/usr/bin/env python3
"""
Batch delete/update the bundled Gemini Enterprise skills.

This is a thin, explicit wrapper around create_skill.py. It targets the two committed Gemini
Enterprise runtime skill bundles:

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
  python3 skill/update_skills.py --live --api-mode public --replace --yes
  python3 skill/update_skills.py --live --api-mode public --delete-only --yes
  python3 skill/update_skills.py --live --api-mode widget --delete-only --yes
  python3 skill/update_skills.py --live --replace --yes --only m365-surface-commander
"""

from __future__ import annotations

import argparse
import hashlib
import os
import re
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


@dataclass(frozen=True)
class UpdatedSkill:
    bundle: SkillBundle
    agent_name: str
    uploaded_zip_path: Path | None = None


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


def _skill_version(bundle: SkillBundle) -> str | None:
    skill_md = HERE / bundle.label / "SKILL.md"
    if not skill_md.exists():
        return None
    text = skill_md.read_text(encoding="utf-8")
    frontmatter = text.split("---", 2)
    source = frontmatter[1] if len(frontmatter) >= 3 and not frontmatter[0].strip() else text
    match = re.search(r"(?m)^\s*version:\s*['\"]?([^'\"\n#]+)", source)
    if not match:
        return None
    return match.group(1).strip()


def _select(names: list[str] | None) -> list[SkillBundle]:
    if not names:
        return list(BUNDLES)
    wanted = set(names)
    bundles = [
        b
        for b in BUNDLES
        if b.agent_id in wanted or b.label in wanted or _normalize_skill_name(b.display_name) in wanted
    ]
    matched = {
        value
        for b in bundles
        for value in (b.agent_id, b.label, _normalize_skill_name(b.display_name))
        if value in wanted
    }
    missing = wanted - matched
    if missing:
        raise SystemExit(f"unknown --only skill(s): {', '.join(sorted(missing))}")
    return bundles


def _normalize_skill_name(value: object) -> str:
    text = str(value or "").strip().lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def _widget_view_name(view: dict) -> str:
    nested_agent = view.get("agent") if isinstance(view.get("agent"), dict) else {}
    return str(view.get("name") or nested_agent.get("name") or "")


def _widget_view_display(view: dict) -> str:
    nested_agent = view.get("agent") if isinstance(view.get("agent"), dict) else {}
    return str(view.get("displayName") or nested_agent.get("displayName") or "")


def _matching_widget_views(bundle: SkillBundle, views: list[dict]) -> list[dict]:
    names = {
        _normalize_skill_name(bundle.label),
        _normalize_skill_name(bundle.display_name),
        _normalize_skill_name(bundle.agent_id),
    }
    matches: list[dict] = []
    seen: set[str] = set()
    for view in views:
        view_name = _widget_view_name(view)
        display_name = _widget_view_display(view)
        normalized = {
            _normalize_skill_name(view_name),
            _normalize_skill_name(display_name),
        }
        if view_name == bundle.agent_id or names & normalized:
            key = view_name or repr(view)
            if key not in seen:
                seen.add(key)
                matches.append(view)
    return matches


def _resolve_widget_targets(
    s,
    cfg: create_skill.LiveConfig,
    widget: create_skill.WidgetConfig,
    bundles: list[SkillBundle],
) -> dict[str, list[str]]:
    views = create_skill.list_widget_agent_views(s, cfg, widget, agent_origin="USER")
    resolved: dict[str, list[str]] = {}
    for bundle in bundles:
        matches = _matching_widget_views(bundle, views)
        resolved[bundle.label] = [_widget_view_name(view) for view in matches if _widget_view_name(view)]
    return resolved


def _write_env_assignments(path: Path, assignments: dict[str, str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = path.read_text(encoding="utf-8").splitlines() if path.exists() else []
    seen: set[str] = set()
    out: list[str] = []
    key_re = re.compile(r"^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=")
    for line in lines:
        match = key_re.match(line)
        if match and match.group(1) in assignments:
            key = match.group(1)
            out.append(f"{key}={assignments[key]}")
            seen.add(key)
        else:
            out.append(line)
    for key, value in assignments.items():
        if key not in seen:
            out.append(f"{key}={value}")
    path.write_text("\n".join(out).rstrip() + "\n", encoding="utf-8")


def _updated_env_assignments(
    cfg: create_skill.LiveConfig,
    widget: create_skill.WidgetConfig | None,
    updated: list[UpdatedSkill],
    *,
    include_bundle_metadata: bool = True,
) -> dict[str, str]:
    assignments = {
        item.bundle.env_key: create_skill.widget_skill_env_value(
            item.bundle.label, cfg, item.agent_name
        )
        for item in updated
    }
    if include_bundle_metadata:
        source_hashes: list[str] = []
        upload_hashes: list[str] = []
        for item in updated:
            bundle = item.bundle
            prefix = bundle.env_key.removesuffix("_SKILL")
            version = _skill_version(bundle)
            if version:
                assignments[f"{prefix}_SKILL_VERSION"] = version
            if bundle.zip_path.exists():
                source_digest = _sha256(bundle.zip_path)
                assignments[f"{prefix}_SKILL_SOURCE_SHA256"] = source_digest
                source_hashes.append(f"{bundle.label}:{source_digest}")
            if item.uploaded_zip_path and item.uploaded_zip_path.exists():
                upload_digest = _sha256(item.uploaded_zip_path)
                assignments[f"{prefix}_SKILL_SHA256"] = upload_digest
                upload_hashes.append(f"{bundle.label}:{upload_digest}")
        if source_hashes:
            combined = hashlib.sha256(
                "\n".join(sorted(source_hashes)).encode("utf-8")
            ).hexdigest()
            assignments["VITE_GE_SKILL_SOURCE_BUNDLE_SET_SHA256"] = combined
        if upload_hashes:
            combined = hashlib.sha256(
                "\n".join(sorted(upload_hashes)).encode("utf-8")
            ).hexdigest()
            assignments["VITE_GE_SKILL_UPLOAD_BUNDLE_SET_SHA256"] = combined
    if widget is not None:
        assignments["VITE_GE_WIDGET_CONFIG_ID"] = widget.config_id
        if widget.server_token:
            assignments["VITE_GE_WIDGET_SERVER_TOKEN"] = widget.server_token
    return assignments


def _updated_from_widget_views(bundles: list[SkillBundle], views: list[dict]) -> list[UpdatedSkill]:
    updated: list[UpdatedSkill] = []
    for bundle in bundles:
        matches = _matching_widget_views(bundle, views)
        names = [_widget_view_name(view) for view in matches if _widget_view_name(view)]
        if len(names) != 1:
            detail = ", ".join(names) if names else "none"
            raise SystemExit(
                f"Cannot write env for {bundle.label}: expected exactly one visible matching "
                f"widget skill, found {len(names)} ({detail}). Use --replace to clean duplicates."
            )
        updated.append(UpdatedSkill(bundle, names[0]))
    return updated


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Delete/update the bundled GE skill zip files.")
    ap.add_argument(
        "--only",
        action="append",
        metavar="SKILL",
        help=(
            "limit to one bundled skill; repeatable. Accepts stable labels such as "
            "m365-surface-commander or the current numeric widget agent id."
        ),
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
        help=(
            "widget matches the Gemini Enterprise web UI; public uses documented OAuth API "
            "(legacy is a backwards-compatible alias for public)"
        ),
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
        "--list",
        action="store_true",
        help="widget mode: list visible skill agent views and stop; no upload/delete",
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
    ap.add_argument(
        "--write-env",
        type=Path,
        help=(
            "write updated non-secret VITE_GE_* skill references to this .env file "
            "(for example packages/web-shell/.env)"
        ),
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
    if args.list and (args.delete_only or args.replace or args.upload_existing or args.create_new):
        raise SystemExit("--list cannot be combined with delete/create/upload modes")
    if args.upload_existing and args.create_new:
        raise SystemExit("--upload-existing and --create-new are mutually exclusive")
    if args.list and args.api_mode != "widget":
        raise SystemExit("--list is implemented for widget mode only")
    if args.api_mode == "widget" and args.share:
        raise SystemExit("--share is only implemented for public API mode")
    if args.api_mode in create_skill.PUBLIC_API_MODES and (args.upload_existing or args.create_new):
        raise SystemExit("--upload-existing/--create-new are widget mode options")

    cfg = create_skill.resolve_live_config()
    widget = create_skill.resolve_widget_config() if args.api_mode == "widget" else None
    bundles = _select(args.only)

    print(
        f"Plan: {'list' if args.list else 'delete' if args.delete_only else 'update'} {len(bundles)} "
        f"Gemini Enterprise skill(s) via {create_skill._api_mode_label(args.api_mode)} API"
    )
    create_skill._print_target(cfg, ",".join(b.agent_id for b in bundles), "BATCH")
    if widget:
        print(f"      widget_config:  {widget.config_id}")
    updated: list[UpdatedSkill] = []
    for bundle in bundles:
        if not bundle.zip_path.exists() and not (args.delete_only or args.list):
            raise SystemExit(f"zip not found: {bundle.zip_path}")
        print(f"\n  skill:       {bundle.agent_id}")
        print(f"  displayName: {bundle.display_name}")
        if not (args.delete_only or args.list):
            print(f"  zip:         {bundle.zip_path}")
            print(f"  sha256:      {_sha256(bundle.zip_path)}")
        if args.replace:
            print("  step:        delete existing agent first")
        if args.delete_only:
            print("  step:        delete only; no upload")
        if args.share:
            print("  step:        share with ALL_USERS after upload")
        if args.api_mode == "widget":
            if args.list:
                print("  step:        list visible widget skill agent views; no upload/delete")
            elif args.replace:
                print("  step:        delete this numeric widget agent, create replacement, upload zip")
            elif args.delete_only:
                print("  step:        delete this numeric widget agent; no replacement/upload")
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

    if args.list:
        assert widget is not None
        s = create_skill.session(cfg, args.api_mode)
        views = create_skill.list_widget_agent_views(s, cfg, widget, agent_origin="USER")
        print(f"\nWidget USER skill agent views: {len(views)}")
        for view in views:
            create_skill._print_widget_agent_view(cfg, view)
        if args.write_env:
            env_assignments = _updated_env_assignments(
                cfg,
                widget,
                _updated_from_widget_views(bundles, views),
                include_bundle_metadata=False,
            )
            _write_env_assignments(args.write_env, env_assignments)
            print(f"\nWrote current web-shell env values to {args.write_env}")
        return 0

    if args.api_mode == "widget" and not (
        args.upload_existing or args.create_new or args.replace or args.delete_only
    ):
        raise SystemExit(
            "Widget live mode requires --replace (with numeric GE_*_AGENT_ID values), "
            "--upload-existing, or --create-new."
        )

    s = create_skill.session(cfg, args.api_mode)
    resolved_widget_targets: dict[str, list[str]] = {}
    if args.api_mode == "widget" and (args.replace or args.delete_only or args.upload_existing):
        assert widget is not None
        resolved_widget_targets = _resolve_widget_targets(s, cfg, widget, bundles)
        print("\nResolved visible widget targets:")
        for bundle in bundles:
            targets = resolved_widget_targets.get(bundle.label, [])
            if targets:
                print(f"  {bundle.label}: {', '.join(targets)}")
                if bundle.agent_id not in targets:
                    print(f"    note: GE env/default id {bundle.agent_id} is stale; using visible target(s)")
            else:
                print(f"  {bundle.label}: no visible match; will fall back to {bundle.agent_id}")

    for bundle in bundles:
        if args.api_mode == "widget":
            assert widget is not None
            print(f"\nUpdate {bundle.agent_id}")
            target_names = resolved_widget_targets.get(bundle.label, []) or [bundle.agent_id]
            if args.delete_only:
                for target_name in target_names:
                    create_skill._print_target(cfg, target_name, "WIDGET DELETE")
                    create_skill.delete_widget_agent(s, cfg, widget, target_name)
                continue
            if args.replace:
                for target_name in target_names:
                    create_skill._print_target(cfg, target_name, "WIDGET DELETE")
                    create_skill.delete_widget_agent(s, cfg, widget, target_name)
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
                if len(target_names) != 1:
                    raise SystemExit(
                        f"--upload-existing for {bundle.label} needs exactly one visible target; "
                        f"found {len(target_names)}: {', '.join(target_names)}"
                    )
                agent_name = target_names[0]
                print(f"  1) use existing agent name {agent_name}")
            print("  2) resumable upload zip (server unpacks)")
            upload_zip_path = create_skill.stamped_skill_zip(
                bundle.zip_path,
                cfg,
                agent_name,
                label=bundle.label,
            )
            print(f"     upload source sha256:  {_sha256(bundle.zip_path)}")
            print(f"     stamped upload sha256: {_sha256(upload_zip_path)}")
            create_skill.upload_zip_resumable(s, cfg, agent_name, upload_zip_path)
            print("  3) verify")
            create_skill.show_widget(s, cfg, widget, agent_name)
            updated.append(UpdatedSkill(bundle, agent_name, upload_zip_path))
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
        upload_zip_path = create_skill.stamped_skill_zip(
            bundle.zip_path,
            cfg,
            bundle.agent_id,
            label=bundle.label,
        )
        print(f"     upload source sha256:  {_sha256(bundle.zip_path)}")
        print(f"     stamped upload sha256: {_sha256(upload_zip_path)}")
        create_skill.upload_zip(s, cfg, bundle.agent_id, upload_zip_path)
        if args.share:
            create_skill._print_target(cfg, bundle.agent_id, "SHARE (ALL_USERS)")
            create_skill.share(s, cfg, bundle.agent_id)
        print("  3) verify")
        create_skill.show(s, cfg, bundle.agent_id)
        updated.append(UpdatedSkill(bundle, bundle.agent_id, upload_zip_path))
    if updated:
        env_assignments = _updated_env_assignments(cfg, widget, updated)
        print("\nUpdated skill references for the add-in env:")
        for item in updated:
            bundle = item.bundle
            print(f"  {bundle.env_key}={env_assignments[bundle.env_key]}")
            prefix = bundle.env_key.removesuffix("_SKILL")
            if f"{prefix}_SKILL_VERSION" in env_assignments:
                print(f"    version: {env_assignments[f'{prefix}_SKILL_VERSION']}")
            if f"{prefix}_SKILL_SOURCE_SHA256" in env_assignments:
                print(f"    source sha256: {env_assignments[f'{prefix}_SKILL_SOURCE_SHA256']}")
            if f"{prefix}_SKILL_SHA256" in env_assignments:
                print(f"    upload sha256: {env_assignments[f'{prefix}_SKILL_SHA256']}")
            print(
                "    mention: "
                + create_skill.widget_skill_mention(
                    bundle.label,
                    env_assignments[bundle.env_key].rsplit("/agents/", 1)[-1],
                )
            )
        if args.write_env:
            _write_env_assignments(args.write_env, env_assignments)
            print(f"\nWrote updated web-shell env values to {args.write_env}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
