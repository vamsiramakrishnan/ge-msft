#!/usr/bin/env python3
"""Generate per-skill resource indexes from Markdown frontmatter."""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SKILLS = ("m365-command-planner", "m365-surface-commander")
SKIP_NAMES = {"SKILL.md", "resource-index.md"}


def split_frontmatter(path: Path) -> tuple[dict[str, str], str] | None:
    text = path.read_text(encoding="utf-8")
    if not text.startswith("---\n"):
        return None
    end = text.find("\n---\n", 4)
    if end == -1:
        return None
    raw = text[4:end]
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
    return data, text[end + 5 :]


def resources(skill_dir: Path) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for md in sorted(skill_dir.rglob("*.md")):
        rel = md.relative_to(skill_dir)
        if md.name in SKIP_NAMES:
            continue
        if rel.parts[0] not in {"references", "patterns", "assets"}:
            continue
        parsed = split_frontmatter(md)
        if parsed is None:
            continue
        frontmatter, _ = parsed
        out.append(
            {
                "path": rel.as_posix(),
                "title": frontmatter.get("title", ""),
                "kind": frontmatter.get("kind", ""),
                "surface": frontmatter.get("surface", ""),
                "workflow": frontmatter.get("workflow", ""),
                "topics": frontmatter.get("topics", ""),
                "load_when": frontmatter.get("load_when", ""),
            },
        )
    return out


def write_index(skill_name: str) -> None:
    skill_dir = ROOT / skill_name
    refs = skill_dir / "references"
    refs.mkdir(exist_ok=True)
    items = resources(skill_dir)

    json_path = refs / "resource-index.json"
    json_path.write_text(json.dumps(items, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    lines = [
        "---",
        "title: Resource Index",
        "kind: index",
        f"skill: {skill_name}",
        "topics: [progressive-disclosure, resources, routing]",
        "load_when: Choosing which bundled reference, pattern, or example to load for a task.",
        "---",
        "",
        "# Resource index",
        "",
        "Generated from resource frontmatter. Load the listed file only when its `load_when` matches the task.",
        "",
        "| Path | Kind | Surface | Workflow | Load when |",
        "| --- | --- | --- | --- | --- |",
    ]
    for item in items:
        lines.append(
            "| "
            + f"[{item['path']}](../{item['path']})"
            + " | "
            + (item["kind"] or "-")
            + " | "
            + (item["surface"] or "-")
            + " | "
            + (item["workflow"] or "-")
            + " | "
            + (item["load_when"] or "-").replace("|", "\\|")
            + " |"
        )
    lines.append("")
    (refs / "resource-index.md").write_text("\n".join(lines), encoding="utf-8")
    print(f"wrote {refs.relative_to(ROOT)}/resource-index.md")
    print(f"wrote {json_path.relative_to(ROOT)}")


def main() -> int:
    for skill in SKILLS:
        write_index(skill)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
