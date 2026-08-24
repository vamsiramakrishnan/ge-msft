#!/usr/bin/env bash

set -euo pipefail

readonly manifest_id="7616958d-7a6e-4750-8737-be42b5f3c23e"
readonly log_name="ge-runtime.txt"
readonly unified_manifest_min_mac_version="16.103"

usage() {
  cat <<'EOF'
Collect Microsoft Office add-in runtime diagnostics on macOS.

Usage:
  scripts/mac-office-addin-debug.sh <word|excel|powerpoint|outlook> [options]

Options:
  --clear-cache  Clear the Office web/Wef cache before relaunching. This can
                 remove every sideloaded Office add-in from the Mac.
  --yes          Skip the --clear-cache confirmation prompt.
  --follow       Continue following the log after printing its last 100 lines.
  -h, --help     Show this help.

Examples:
  scripts/mac-office-addin-debug.sh word
  scripts/mac-office-addin-debug.sh excel --follow
  scripts/mac-office-addin-debug.sh word --clear-cache
EOF
}

host=""
clear_cache=false
assume_yes=false
follow=false

while (($# > 0)); do
  case "$1" in
    word | excel | powerpoint | outlook)
      if [[ -n "$host" ]]; then
        printf 'Only one Office host can be debugged at a time.\n' >&2
        exit 2
      fi
      host="$1"
      ;;
    --clear-cache)
      clear_cache=true
      ;;
    --yes)
      assume_yes=true
      ;;
    --follow)
      follow=true
      ;;
    -h | --help)
      usage
      exit 0
      ;;
    *)
      printf 'Unknown argument: %s\n\n' "$1" >&2
      usage >&2
      exit 2
      ;;
  esac
  shift
done

if [[ -z "$host" ]]; then
  usage >&2
  exit 2
fi

if [[ "$(uname -s)" != "Darwin" ]]; then
  printf 'This helper must run on the Mac that hosts Microsoft Office.\n' >&2
  exit 2
fi

case "$host" in
  word)
    bundle_id="com.microsoft.Word"
    app_name="Microsoft Word"
    process_name="Microsoft Word"
    ;;
  excel)
    bundle_id="com.microsoft.Excel"
    app_name="Microsoft Excel"
    process_name="Microsoft Excel"
    ;;
  powerpoint)
    bundle_id="com.microsoft.Powerpoint"
    app_name="Microsoft PowerPoint"
    process_name="Microsoft PowerPoint"
    ;;
  outlook)
    bundle_id="com.microsoft.Outlook"
    app_name="Microsoft Outlook"
    process_name="Microsoft Outlook"
    ;;
esac

for command_name in awk defaults osascript pgrep open tail; do
  if ! command -v "$command_name" >/dev/null 2>&1; then
    printf 'Required command is unavailable: %s\n' "$command_name" >&2
    exit 1
  fi
done

readonly app_bundle="/Applications/$app_name.app"
readonly info_plist="$app_bundle/Contents/Info.plist"

if [[ ! -f "$info_plist" ]]; then
  printf '%s is not installed at the expected path: %s\n' "$app_name" "$app_bundle" >&2
  exit 1
fi

office_version="$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "$info_plist" 2>/dev/null || true)"
office_build="$(/usr/libexec/PlistBuddy -c 'Print :CFBundleVersion' "$info_plist" 2>/dev/null || true)"

printf '%s version: %s (build %s)\n' \
  "$app_name" "${office_version:-unknown}" "${office_build:-unknown}"

version_at_least() {
  awk -v actual="$1" -v minimum="$2" 'BEGIN {
    split(actual, actual_parts, ".")
    split(minimum, minimum_parts, ".")
    if ((actual_parts[1] + 0) > (minimum_parts[1] + 0)) exit 0
    if ((actual_parts[1] + 0) < (minimum_parts[1] + 0)) exit 1
    exit !((actual_parts[2] + 0) >= (minimum_parts[2] + 0))
  }'
}

if [[ "$host" == word || "$host" == excel || "$host" == powerpoint ]]; then
  if [[ -n "$office_version" ]] && ! version_at_least "$office_version" "$unified_manifest_min_mac_version"; then
    printf 'Unsupported Office build: unified-manifest sideloading on Mac requires version %s or later.\n' \
      "$unified_manifest_min_mac_version" >&2
    printf 'Update Office with Microsoft AutoUpdate, then rerun this helper.\n' >&2
    exit 1
  fi
elif [[ "$host" == outlook ]]; then
  printf 'Warning: Microsoft does not currently support unified-manifest sideloading in Outlook on Mac.\n' >&2
fi

readonly container_data="$HOME/Library/Containers/$bundle_id/Data"
readonly log_path="$container_data/$log_name"

printf 'Enabling Office runtime logging and Web Inspector for %s...\n' "$app_name"
defaults write "$bundle_id" CEFRuntimeLoggingFile -string "$log_name"
defaults write "$bundle_id" OfficeWebAddinDeveloperExtras -bool true

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
    printf '%s is still running. Save open documents and quit it from Activity Monitor, then rerun this script.\n' "$app_name" >&2
    exit 1
  fi
fi

if [[ -f "$log_path" ]]; then
  archive_path="$container_data/ge-runtime-$(date -u +%Y%m%dT%H%M%SZ).txt"
  mv "$log_path" "$archive_path"
  printf 'Archived the previous log: %s\n' "$archive_path"
fi

if [[ "$clear_cache" == true ]]; then
  if ! command -v npx >/dev/null 2>&1; then
    printf 'npx is required for --clear-cache but was not found.\n' >&2
    exit 1
  fi

  if [[ "$assume_yes" != true ]]; then
    printf 'Warning: clearing the Office cache can remove every sideloaded Office add-in on this Mac. Continue? [y/N] '
    read -r confirmation
    case "$confirmation" in
      y | Y | yes | YES) ;;
      *)
        printf 'Cache clear cancelled.\n'
        exit 0
        ;;
    esac
  fi

  npx office-addin-cache clear
fi

printf 'Launching %s...\n' "$app_name"
open -b "$bundle_id"

printf '\nOpen "Gemini Enterprise Dev" in %s and reproduce the problem.\n' "$app_name"
printf 'When the pane has loaded or failed, press Return here to inspect the log.\n'
read -r

if [[ ! -f "$log_path" ]]; then
  printf 'Office did not create %s.\n' "$log_path" >&2
  printf 'Logging preference: '
  defaults read "$bundle_id" CEFRuntimeLoggingFile 2>/dev/null || true
  printf 'Confirm that Office was installed from Office.com and that the add-in was opened once.\n' >&2
  exit 1
fi

printf '\nLast 100 runtime-log lines:\n'
printf '%s\n' '------------------------------------------------------------------------'
tail -n 100 "$log_path"
printf '%s\n' '------------------------------------------------------------------------'
printf 'Manifest ID to search for: %s\n' "$manifest_id"
printf 'Full log: %s\n' "$log_path"
printf 'Redact tokens, Authorization headers, and OAuth code/state values before sharing output.\n'

if [[ "$follow" == true ]]; then
  printf '\nFollowing the log. Press Ctrl-C to stop.\n'
  tail -F "$log_path"
fi
