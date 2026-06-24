#!/usr/bin/env bash
# (Re)build a skill bundle zip from a skill directory.
# SKILL.md is placed at the ARCHIVE ROOT (Gemini Enterprise requires SKILL.md in the package root).
#
# Usage:
#   ./build_zip.sh                       # builds m365-surface-commander.zip (default)
#   ./build_zip.sh m365-command-planner  # builds m365-command-planner.zip
set -euo pipefail
cd "$(dirname "$0")"

SKILL_DIR="${1:-m365-surface-commander}"
SKILL_DIR="${SKILL_DIR%/}"
[ -f "$SKILL_DIR/SKILL.md" ] || { echo "no SKILL.md in '$SKILL_DIR'" >&2; exit 1; }
OUT="$SKILL_DIR.zip"

# include only the bundle subdirs that exist (the planner has no assets/, etc.)
INCLUDE=(SKILL.md)
for d in references scripts assets; do
  [ -d "$SKILL_DIR/$d" ] && INCLUDE+=("$d")
done

rm -f "$OUT"
( cd "$SKILL_DIR" && zip -r -X "../$OUT" "${INCLUDE[@]}" \
    -x '*.pyc' -x '*__pycache__*' >/dev/null )
echo "built $OUT"
unzip -l "$OUT"
