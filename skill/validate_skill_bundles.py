#!/usr/bin/env python3
"""Validate local Gemini Enterprise skill bundles.

This is intentionally dependency-free: it checks the progressive-disclosure structure that keeps
SKILL.md concise while making references/assets discoverable.
"""

from __future__ import annotations

import argparse
import importlib.util
import re
import sys
from bundle import validate_archive
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DEFAULT_SKILLS = (
    "m365-command-planner",
    "m365-surface-commander",
    "m365-release-operator",
)
# The platform ingests SKILL.md as the always-on instruction. Keep that entrypoint small enough to
# reach the first useful token quickly; put exact syntax, examples, and long-tail behavior in routed
# references instead. All current bundles fit comfortably under both limits.
MAX_SKILL_LINES = 200
MAX_SKILL_BYTES = 10_000

REQUIRED_SKILL_FRONTMATTER = ("name", "description")
REQUIRED_RESOURCE_FRONTMATTER = ("title", "kind", "skill", "topics", "load_when")
ALLOWED_RESOURCE_KINDS = {"example", "generated-reference", "index", "pattern", "reference"}
ALLOWED_SURFACES = {"excel", "onenote", "outlook", "powerpoint", "teams", "word"}
ALLOWED_WORKFLOWS = {"cross-surface", "single-surface"}
MARKDOWN_LINK = re.compile(r"\[[^\]]+\]\(([^)]+)\)")
PLAN_BLOCK = re.compile(r"```plan[^\S\n]*\r?\n([\s\S]*?)```", re.IGNORECASE)
CMD_BLOCK = re.compile(r"```cmd[^\S\n]*\r?\n([\s\S]*?)```", re.IGNORECASE)


def split_frontmatter(path: Path) -> tuple[dict[str, str], str] | None:
    text = path.read_text(encoding="utf-8")
    if not text.startswith("---\n"):
        return None
    end = text.find("\n---\n", 4)
    if end == -1:
        return None
    raw = text[4:end]
    body = text[end + 5 :]
    data: dict[str, str] = {}
    current_key: str | None = None
    for line in raw.splitlines():
        if not line.strip():
            continue
        if line.startswith((" ", "\t")):
            if current_key:
                data[current_key] = f"{data[current_key]} {line.strip()}".strip()
            continue
        if ":" in line:
            key, value = line.split(":", 1)
            current_key = key.strip()
            data[current_key] = value.strip()
    return data, body


def validate_links(path: Path) -> list[str]:
    errors: list[str] = []
    text = path.read_text(encoding="utf-8")
    for match in MARKDOWN_LINK.finditer(text):
        target = match.group(1).strip()
        if not target or target.startswith(("http://", "https://", "#", "mailto:")):
            continue
        file_part = target.split("#", 1)[0]
        if not file_part:
            continue
        resolved = (path.parent / file_part).resolve()
        if not resolved.exists():
            errors.append(f"{path.relative_to(ROOT)} links to missing {target}")
    return errors


def parse_inline_list(value: str) -> list[str]:
    stripped = value.strip()
    if not stripped.startswith("[") or not stripped.endswith("]"):
        return []
    items = [item.strip().strip("\"'") for item in stripped[1:-1].split(",")]
    return [item for item in items if item]


def validate_resource_frontmatter(
    skill_dir: Path,
    path: Path,
    frontmatter: dict[str, str],
    body: str,
) -> list[str]:
    errors: list[str] = []
    rel = path.relative_to(ROOT)
    skill_rel = path.relative_to(skill_dir)
    folder = skill_rel.parts[0] if len(skill_rel.parts) > 1 else ""

    kind = frontmatter.get("kind", "")
    if kind not in ALLOWED_RESOURCE_KINDS:
        errors.append(
            f"{rel} frontmatter kind must be one of {', '.join(sorted(ALLOWED_RESOURCE_KINDS))}",
        )

    expected_by_folder = {
        "assets": {"example"},
        "patterns": {"pattern"},
        "references": {"generated-reference", "index", "reference"},
    }
    expected_kinds = expected_by_folder.get(folder)
    if expected_kinds and kind not in expected_kinds:
        errors.append(
            f"{rel} frontmatter kind {kind!r} does not match {folder}/; expected "
            f"{', '.join(sorted(expected_kinds))}",
        )

    topics = parse_inline_list(frontmatter.get("topics", ""))
    if not topics:
        errors.append(f"{rel} frontmatter topics must be a non-empty inline list")

    load_when = frontmatter.get("load_when", "").strip()
    if len(load_when) < 20:
        errors.append(f"{rel} frontmatter load_when must be a useful routing sentence")

    surface = frontmatter.get("surface")
    if surface and surface not in ALLOWED_SURFACES:
        errors.append(
            f"{rel} frontmatter surface must be one of {', '.join(sorted(ALLOWED_SURFACES))}",
        )

    workflow = frontmatter.get("workflow")
    if workflow and workflow not in ALLOWED_WORKFLOWS:
        errors.append(
            f"{rel} frontmatter workflow must be one of {', '.join(sorted(ALLOWED_WORKFLOWS))}",
        )

    if kind == "generated-reference" and "<!-- GENERATED" not in body:
        errors.append(f"{rel} generated-reference must contain the generated-file marker")

    return errors


def validate_skill(skill_dir: Path) -> list[str]:
    errors: list[str] = []
    skill_md = skill_dir / "SKILL.md"
    if not skill_md.exists():
        return [f"{skill_dir.name}: missing SKILL.md"]

    fm = split_frontmatter(skill_md)
    if fm is None:
        errors.append(f"{skill_dir.name}/SKILL.md missing YAML frontmatter")
    else:
        frontmatter, _ = fm
        for key in REQUIRED_SKILL_FRONTMATTER:
            if not frontmatter.get(key):
                errors.append(f"{skill_dir.name}/SKILL.md missing frontmatter field {key}")

    skill_text = skill_md.read_text(encoding="utf-8")
    line_count = len(skill_text.splitlines())
    if line_count > MAX_SKILL_LINES:
        errors.append(f"{skill_dir.name}/SKILL.md has {line_count} lines; max is {MAX_SKILL_LINES}")
    byte_count = len(skill_text.encode("utf-8"))
    if byte_count > MAX_SKILL_BYTES:
        errors.append(
            f"{skill_dir.name}/SKILL.md has {byte_count} bytes; max is {MAX_SKILL_BYTES}; "
            "move exact detail into a routed reference",
        )
    if (skill_dir / "references").is_dir():
        if "references/resource-index.md" not in skill_text:
            errors.append(
                f"{skill_dir.name}/SKILL.md must link references/resource-index.md for "
                "progressive disclosure",
            )
        if "load" not in skill_text.lower() or "only" not in skill_text.lower():
            errors.append(
                f"{skill_dir.name}/SKILL.md must explain selective reference loading",
            )

    for md in sorted(skill_dir.rglob("*.md")):
        errors.extend(validate_links(md))
        if md.name == "SKILL.md":
            continue
        fm = split_frontmatter(md)
        rel = md.relative_to(ROOT)
        if fm is None:
            errors.append(f"{rel} missing YAML frontmatter")
            continue
        frontmatter, body = fm
        for key in REQUIRED_RESOURCE_FRONTMATTER:
            if not frontmatter.get(key):
                errors.append(f"{rel} missing frontmatter field {key}")
        if frontmatter.get("skill") != skill_dir.name:
            errors.append(f"{rel} frontmatter skill must be {skill_dir.name}")
        errors.extend(validate_resource_frontmatter(skill_dir, md, frontmatter, body))

    return errors


def validate_zip(skill_name: str) -> list[str]:
    return validate_archive(ROOT / skill_name, ROOT / f"{skill_name}.zip")


def import_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"could not load {path}")
    module = importlib.util.module_from_spec(spec)
    # Resolve each standalone parser's bundled helper, without accidentally reusing the other
    # skill's previously imported copy. A missing helper must remain a packaging failure.
    previous_helper = sys.modules.pop("language_manifest", None)
    sys.path.insert(0, str(path.parent))
    try:
        spec.loader.exec_module(module)
    finally:
        sys.path.pop(0)
        sys.modules.pop("language_manifest", None)
        if previous_helper is not None:
            sys.modules["language_manifest"] = previous_helper
    return module


def validate_planner_fixtures(skill_dir: Path) -> list[str]:
    errors: list[str] = []
    parser = import_module(skill_dir / "scripts" / "parse_plan.py", "m365_parse_plan")
    for md in sorted((skill_dir / "assets" / "example-plans").glob("*.md")):
        text = md.read_text(encoding="utf-8")
        blocks = PLAN_BLOCK.findall(text)
        if not blocks:
            errors.append(f"{md.relative_to(ROOT)} has no ```plan block")
            continue
        for index, block in enumerate(blocks, start=1):
            parsed = parser.parse_plan(f"```plan\n{block.strip()}\n```")
            if parsed["errors"]:
                errors.append(
                    f"{md.relative_to(ROOT)} plan block {index} has errors: {parsed['errors']}",
                )
    return errors


def validate_commander_fixtures(skill_dir: Path) -> list[str]:
    errors: list[str] = []
    parser = import_module(
        skill_dir / "scripts" / "parse_commands.py",
        "m365_parse_commands",
    )
    fixture_roots = [skill_dir / "assets" / "example-sessions", skill_dir / "patterns"]
    for root in fixture_roots:
        for md in sorted(root.glob("*.md")):
            text = md.read_text(encoding="utf-8")
            blocks = CMD_BLOCK.findall(text)
            if not blocks:
                continue
            for index, block in enumerate(blocks, start=1):
                parsed = parser.parse_block(f"```cmd\n{block.strip()}\n```")
                bad = [cmd["error"] for cmd in parsed["commands"] if "error" in cmd]
                if bad:
                    errors.append(
                        f"{md.relative_to(ROOT)} cmd block {index} has parse errors: {bad}",
                    )
    return errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("skills", nargs="*", default=list(DEFAULT_SKILLS))
    parser.add_argument("--check-zip", action="store_true", help="also validate built zip contents")
    args = parser.parse_args()

    errors: list[str] = []
    for name in args.skills:
        skill_dir = ROOT / name
        if not skill_dir.exists():
            errors.append(f"missing skill directory {name}")
            continue
        errors.extend(validate_skill(skill_dir))
        if name == "m365-command-planner":
            errors.extend(validate_planner_fixtures(skill_dir))
        if name == "m365-surface-commander":
            errors.extend(validate_commander_fixtures(skill_dir))
        if args.check_zip:
            errors.extend(validate_zip(name))

    if errors:
        for error in errors:
            print(f"ERROR: {error}", file=sys.stderr)
        return 1
    print("SKILL BUNDLE VALIDATION OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
