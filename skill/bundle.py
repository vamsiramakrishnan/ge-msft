"""One archive inventory for deterministic skill builds and exact source validation."""
from __future__ import annotations

import argparse
from pathlib import Path, PurePosixPath
import zipfile
import stat

ROOT = Path(__file__).resolve().parent
INCLUDED = {"references", "patterns", "scripts", "assets"}


def inventory(skill_dir: Path) -> dict[str, bytes]:
    if not (skill_dir / "SKILL.md").is_file():
        raise ValueError(f"Missing {skill_dir}/SKILL.md")
    if (skill_dir / "SKILL.md").is_symlink():
        raise ValueError("SKILL.md cannot be a symlink")
    files = {"SKILL.md": (skill_dir / "SKILL.md").read_bytes()}
    for folder in sorted(INCLUDED):
        if (skill_dir / folder).is_symlink():
            raise ValueError("Skill resource directories cannot be symlinks")
        for path in sorted((skill_dir / folder).rglob("*")):
            if "__pycache__" in path.parts or path.suffix == ".pyc":
                continue
            if path.is_symlink():
                raise ValueError(f"Skill archives cannot contain symlinks: {path}")
            if path.is_file():
                files[path.relative_to(skill_dir).as_posix()] = path.read_bytes()
    return dict(sorted(files.items()))


def build(skill_dir: Path, output: Path) -> None:
    files = inventory(skill_dir)
    temporary = output.with_suffix(output.suffix + ".tmp")
    try:
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            for name, content in files.items():
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.create_system = 3
                info.external_attr = 0o100644 << 16
                info.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(info, content)
        temporary.replace(output)
    finally:
        temporary.unlink(missing_ok=True)


def validate_archive(skill_dir: Path, archive_path: Path) -> list[str]:
    if not archive_path.is_file():
        return [f"Missing built skill archive {archive_path.name}"]
    try:
        expected = inventory(skill_dir)
        with zipfile.ZipFile(archive_path) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                return [f"{archive_path.name} contains duplicate entries"]
            for info in archive.infolist():
                if stat.S_ISLNK(info.external_attr >> 16):
                    return [f"{archive_path.name} contains a symbolic link"]
            for name in names:
                if name.startswith('/') or '..' in PurePosixPath(name).parts:
                    return [f"{archive_path.name} contains an unsafe entry"]
            if set(names) != set(expected):
                return [f"{archive_path.name} file inventory differs from current source"]
            if any(archive.getinfo(name).file_size != len(content) for name, content in expected.items()):
                return [f"{archive_path.name} file sizes differ from current source"]
            if any(archive.read(name) != content for name, content in expected.items()):
                return [f"{archive_path.name} content differs from current source; rebuild the skill"]
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as error:
        return [f"Invalid {archive_path.name}: {error}"]
    return []


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('skill', nargs='?', default='m365-surface-commander')
    args = parser.parse_args()
    args.skill = args.skill.rstrip("/")
    # Preserve the wrapper's skill-name interface; never write outside the skill root.
    if Path(args.skill).name != args.skill:
        parser.error('Use a skill directory name, not a path')
    directory = ROOT / args.skill
    output = ROOT / f'{args.skill}.zip'
    build(directory, output)
    print(f'Built {output.name}: {len(inventory(directory))} files')


if __name__ == '__main__':
    main()
