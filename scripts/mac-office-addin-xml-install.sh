#!/usr/bin/env bash

set -euo pipefail

usage() {
  cat <<'EOF'
Install a generated add-in-only XML manifest into an Office host on macOS.

Usage:
  mac-office-addin-xml-install.sh <word|excel|powerpoint> <manifest.xml>

This is a local desktop bypass for the Microsoft work-account Add-ins catalog
spinner. It does not remove or modify the tenant-installed unified package.
EOF
}

if (($# != 2)); then
  usage >&2
  exit 2
fi

host="$1"
manifest_path="$2"

if [[ "$(uname -s)" != "Darwin" ]]; then
  printf 'This helper must run on the Mac that hosts Microsoft Office.\n' >&2
  exit 2
fi

case "$host" in
  word)
    bundle_id="com.microsoft.Word"
    app_name="Microsoft Word"
    process_name="Microsoft Word"
    expected_host='Document'
    ;;
  excel)
    bundle_id="com.microsoft.Excel"
    app_name="Microsoft Excel"
    process_name="Microsoft Excel"
    expected_host='Workbook'
    ;;
  powerpoint)
    bundle_id="com.microsoft.Powerpoint"
    app_name="Microsoft PowerPoint"
    process_name="Microsoft PowerPoint"
    expected_host='Presentation'
    ;;
  *)
    usage >&2
    exit 2
    ;;
esac

if [[ ! -f "$manifest_path" ]]; then
  printf 'Manifest not found: %s\n' "$manifest_path" >&2
  exit 1
fi

if ! grep -q '<OfficeApp ' "$manifest_path" ||
  ! grep -q "<Host Name=\"$expected_host\"" "$manifest_path" ||
  ! grep -q '<Control xsi:type="Button" id="openGeminiBtn">' "$manifest_path"; then
  printf 'Manifest is not the generated Gemini Enterprise %s desktop manifest: %s\n' \
    "$host" "$manifest_path" >&2
  exit 1
fi

for command_name in cp date grep mkdir mv open osascript pgrep sleep; do
  if ! command -v "$command_name" >/dev/null 2>&1; then
    printf 'Required command is unavailable: %s\n' "$command_name" >&2
    exit 1
  fi
done

if pgrep -x "$process_name" >/dev/null 2>&1; then
  printf 'Asking %s to quit cleanly...\n' "$app_name"
  osascript -e "tell application \"$app_name\" to quit"

  for ((attempt = 0; attempt < 30; attempt += 1)); do
    if ! pgrep -x "$process_name" >/dev/null 2>&1; then
      break
    fi
    sleep 1
  done

  if pgrep -x "$process_name" >/dev/null 2>&1; then
    printf '%s is still running. Save open files, quit it, and rerun this helper.\n' \
      "$app_name" >&2
    exit 1
  fi
fi

wef_dir="$HOME/Library/Containers/$bundle_id/Data/Documents/wef"
installed_manifest="$wef_dir/gemini-enterprise-$host.manifest.xml"
mkdir -p "$wef_dir"

if [[ -f "$installed_manifest" ]]; then
  backup_manifest="$installed_manifest.$(date -u +%Y%m%dT%H%M%SZ).bak"
  mv "$installed_manifest" "$backup_manifest"
  printf 'Backed up the previous local manifest: %s\n' "$backup_manifest"
fi

cp "$manifest_path" "$installed_manifest"
printf 'Installed local desktop manifest: %s\n' "$installed_manifest"
printf 'The tenant unified package was not changed.\n'

open -b "$bundle_id"
printf 'Look for Gemini Enterprise on the Home ribbon in %s.\n' "$app_name"
