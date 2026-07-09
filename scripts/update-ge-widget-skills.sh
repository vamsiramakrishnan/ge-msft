#!/usr/bin/env bash
# Guided Gemini Enterprise widget-skill updater for the dev app.
#
# What this does:
#   1. Rebuild and validate the local skill bundles.
#   2. Open the Gemini Enterprise app URL so you can copy one authenticated
#      content-discoveryengine.googleapis.com request as cURL/HAR from DevTools.
#   3. Extract the short-lived widget bearer token/config into /tmp.
#   4. List current private skills.
#   5. Delete and recreate the selected skills, upload the latest zips, and verify.
#
# The script never reads browser cookies or XSRF state. It only consumes a cURL/HAR
# file that you explicitly save locally.

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

GE_PROJECT="${GE_PROJECT:-saib-ai-playground}"
GE_PROJECT_NUMBER="${GE_PROJECT_NUMBER:-288406675721}"
GE_LOCATION="${GE_LOCATION:-global}"
GE_ENGINE="${GE_ENGINE:-ge-msft-plugin-test_1782382759735}"
GE_WIDGET_CONFIG_ID="${GE_WIDGET_CONFIG_ID:-cd8248bf-0b65-487d-9a81-fdd48f3912e7}"
GE_WIDGET_SERVER_TOKEN="${GE_WIDGET_SERVER_TOKEN:-CAMSAh0H}"
GE_COMMAND_PLANNER_AGENT_ID="${GE_COMMAND_PLANNER_AGENT_ID:-m365-command-planner}"
GE_SURFACE_COMMANDER_AGENT_ID="${GE_SURFACE_COMMANDER_AGENT_ID:-m365-surface-commander}"

TOKEN_FILE="${GE_WIDGET_BEARER_TOKEN_FILE:-/tmp/ge-widget-token}"
ENV_FILE="/tmp/ge-widget.env"
WEB_SHELL_ENV="${GE_WEB_SHELL_ENV:-${ROOT_DIR}/packages/web-shell/.env}"
TOKEN_MIN_TTL_SECONDS="${GE_WIDGET_MIN_TTL_SECONDS:-120}"
CREDENTIALS_FILE=""
PASTE_CREDENTIALS=0
FORCE_TOKEN_REFRESH=0
OPEN_BROWSER=1
SKIP_REBUILD=0
DRY_RUN=0
LIST_ONLY=0
WRITE_WEB_SHELL_ENV=1
YES=0

usage() {
  cat <<'EOF'
Usage:
  scripts/update-ge-widget-skills.sh [options]

Options:
  --credentials-file PATH   DevTools-exported cURL/HAR/text request containing a widget Bearer token.
                            Use "-" to read the request from stdin.
  --paste-curl              Prompt for a pasted multi-line cURL/HAR block, ending with __GE_WIDGET_CURL_END__.
  --env-file PATH           Sourceable temp env output. Default: /tmp/ge-widget.env
  --token-file PATH         Temp token file. Default: /tmp/ge-widget-token
  --force-token-refresh     Ignore a still-valid token file and ask for a fresh widget cURL/HAR.
  --min-token-ttl SECONDS   Reuse the token only when at least this many seconds remain.
                            Default: 120. Also configurable with GE_WIDGET_MIN_TTL_SECONDS.
  --web-shell-env PATH      Web-shell .env file to update after replacement.
                            Default: packages/web-shell/.env
  --no-write-web-shell-env  Do not update the web-shell .env after replacement.
  --no-open                 Do not try to open the Gemini Enterprise app URL.
  --skip-rebuild            Skip Bun/skill validation and zip rebuild.
  --dry-run                 Build/extract/list, then print replace plan without --live.
  --list-only               Build/extract/list current skills and stop.
  --yes                     Do not prompt before destructive replace.
  -h, --help                Show this help.

Environment overrides:
  GE_PROJECT, GE_PROJECT_NUMBER, GE_LOCATION, GE_ENGINE, GE_WIDGET_CONFIG_ID,
  GE_WIDGET_SERVER_TOKEN, GE_COMMAND_PLANNER_AGENT_ID, GE_SURFACE_COMMANDER_AGENT_ID,
  GE_WEB_SHELL_ENV, GE_WIDGET_MIN_TTL_SECONDS.

Typical flow:
  1. Run this script.
  2. In the opened Gemini Enterprise tab, open DevTools > Network.
  3. Trigger any Gemini Enterprise request, filter for content-discoveryengine.googleapis.com.
  4. Copy a widget request as cURL or export HAR.
  5. Paste the saved file path, one-line cURL, or type PASTE to paste a multi-line block.
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --credentials-file)
      CREDENTIALS_FILE="${2:-}"
      shift 2
      ;;
    --paste-curl)
      PASTE_CREDENTIALS=1
      shift
      ;;
    --env-file)
      ENV_FILE="${2:-}"
      shift 2
      ;;
    --token-file)
      TOKEN_FILE="${2:-}"
      shift 2
      ;;
    --force-token-refresh)
      FORCE_TOKEN_REFRESH=1
      shift
      ;;
    --min-token-ttl)
      TOKEN_MIN_TTL_SECONDS="${2:-}"
      shift 2
      ;;
    --web-shell-env)
      WEB_SHELL_ENV="${2:-}"
      shift 2
      ;;
    --no-write-web-shell-env)
      WRITE_WEB_SHELL_ENV=0
      shift
      ;;
    --no-open)
      OPEN_BROWSER=0
      shift
      ;;
    --skip-rebuild)
      SKIP_REBUILD=1
      shift
      ;;
    --dry-run)
      DRY_RUN=1
      shift
      ;;
    --list-only)
      LIST_ONLY=1
      shift
      ;;
    --yes)
      YES=1
      shift
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown option: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

export GE_PROJECT
export GE_PROJECT_NUMBER
export GE_LOCATION
export GE_ENGINE
export GE_WIDGET_CONFIG_ID
export GE_WIDGET_SERVER_TOKEN
export GE_COMMAND_PLANNER_AGENT_ID
export GE_SURFACE_COMMANDER_AGENT_ID

APP_URL="https://vertexaisearch.cloud.google/home/cid/${GE_WIDGET_CONFIG_ID}"

section() {
  printf '\n==> %s\n' "$1"
}

source_widget_env_if_present() {
  if [[ -f "${ENV_FILE}" ]]; then
    # shellcheck disable=SC1090
    source "${ENV_FILE}"
  fi
  export GE_WIDGET_BEARER_TOKEN_FILE="${GE_WIDGET_BEARER_TOKEN_FILE:-${TOKEN_FILE}}"
}

widget_token_status() {
  python3 - "${ROOT_DIR}" "${TOKEN_FILE}" "${TOKEN_MIN_TTL_SECONDS}" <<'PY'
import sys
import time
from pathlib import Path

root = Path(sys.argv[1])
token_path = Path(sys.argv[2])
min_ttl = int(float(sys.argv[3]))
sys.path.insert(0, str(root / "skill"))

import create_skill  # noqa: E402

if not token_path.exists():
    print(f"missing token file: {token_path}")
    raise SystemExit(1)

token = token_path.read_text(encoding="utf-8").strip().removeprefix("Bearer ").strip()
if not token:
    print(f"empty token file: {token_path}")
    raise SystemExit(1)

try:
    payload = create_skill._jwt_payload(token)
    create_skill._validate_widget_bearer_token(token)
except BaseException as exc:
    print(f"invalid token: {exc}")
    raise SystemExit(1)

exp = payload.get("exp")
if not isinstance(exp, (int, float)):
    print("valid widget token: no exp claim")
    raise SystemExit(0)

remaining = int(exp - time.time())
if remaining < min_ttl:
    print(f"widget token has only {remaining}s remaining; refresh required")
    raise SystemExit(2)

minutes, seconds = divmod(max(0, remaining), 60)
print(f"valid widget token: {minutes}m {seconds}s remaining")
PY
}

reuse_widget_token_if_valid() {
  if [[ "${FORCE_TOKEN_REFRESH}" == "1" || -n "${CREDENTIALS_FILE}" || "${PASTE_CREDENTIALS}" == "1" ]]; then
    return 1
  fi

  local status
  if status="$(widget_token_status 2>&1)"; then
    source_widget_env_if_present
    section "Reuse widget token"
    printf '  %s\n' "${status}"
    printf '  token_file: %s\n' "${TOKEN_FILE}"
    return 0
  fi

  section "Widget token refresh required"
  printf '  %s\n' "${status}"
  return 1
}

open_app_url() {
  section "Open Gemini Enterprise"
  echo "App URL: ${APP_URL}"
  if [[ "${OPEN_BROWSER}" != "1" ]]; then
    return 0
  fi
  if python3 - "${APP_URL}" <<'PY'
import sys
import webbrowser

sys.exit(0 if webbrowser.open(sys.argv[1], new=2) else 1)
PY
  then
    echo "Browser open requested."
  else
    echo "Could not open a browser from this shell. Copy the URL above."
  fi
}

rebuild_and_validate() {
  if [[ "${SKIP_REBUILD}" == "1" ]]; then
    section "Skip rebuild"
    return 0
  fi
  section "Validate generated skill artifacts"
  cd "${ROOT_DIR}"
  bun run skills:check
  python3 skill/validate_skill_bundles.py

  section "Rebuild skill zips"
  cd "${ROOT_DIR}/skill"
  ./build_zip.sh m365-command-planner
  ./build_zip.sh m365-surface-commander
  cd "${ROOT_DIR}"
}

file_sha256() {
  python3 - "$1" <<'PY'
import hashlib
import sys
from pathlib import Path

h = hashlib.sha256()
with Path(sys.argv[1]).open("rb") as fh:
    for chunk in iter(lambda: fh.read(1024 * 1024), b""):
        h.update(chunk)
print(h.hexdigest())
PY
}

env_value() {
  local key="$1"
  if [[ ! -f "${WEB_SHELL_ENV}" ]]; then
    return 0
  fi
  python3 - "${WEB_SHELL_ENV}" "${key}" <<'PY'
import re
import sys
from pathlib import Path

path = Path(sys.argv[1])
key = sys.argv[2]
pattern = re.compile(r"^\s*(?:export\s+)?" + re.escape(key) + r"\s*=\s*(.*)\s*$")
value = ""
for line in path.read_text(encoding="utf-8").splitlines():
    match = pattern.match(line)
    if match:
        value = match.group(1).strip().strip("'\"")
if value:
    print(value)
PY
}

report_bundle_drift() {
  section "Skill bundle drift"
  local status="clean"
  local label zip key current recorded version
  for item in \
    "m365-command-planner:${ROOT_DIR}/skill/m365-command-planner.zip:VITE_GE_COMMAND_PLANNER_SKILL_SOURCE_SHA256:VITE_GE_COMMAND_PLANNER_SKILL_VERSION" \
    "m365-surface-commander:${ROOT_DIR}/skill/m365-surface-commander.zip:VITE_GE_SURFACE_COMMANDER_SKILL_SOURCE_SHA256:VITE_GE_SURFACE_COMMANDER_SKILL_VERSION"
  do
    IFS=: read -r label zip key version_key <<<"${item}"
    current="$(file_sha256 "${zip}")"
    recorded="$(env_value "${key}")"
    version="$(env_value "${version_key}")"
    if [[ -z "${recorded}" ]]; then
      status="unknown"
      printf '%s\n' "  ${label}: no recorded source hash yet (upload will stamp one)"
    elif [[ "${recorded}" == "${current}" ]]; then
      printf '%s\n' "  ${label}: source hash matches ${current:0:12}${version:+, version ${version}}"
    else
      status="dirty"
      printf '%s\n' "  ${label}: source hash changed"
      printf '%s\n' "    recorded: ${recorded}"
      printf '%s\n' "    current:  ${current}"
    fi
  done
  case "${status}" in
    clean) echo "Local skill zips match the recorded uploaded source hashes." ;;
    unknown) echo "No complete recorded source hash set yet; run replacement to establish provenance." ;;
    dirty) echo "At least one skill zip changed; replacement is needed to sync Gemini Enterprise." ;;
  esac
}

credentials_file() {
  local tmp
  tmp="$(mktemp /tmp/ge-widget-request.XXXXXX)"

  if [[ -n "${CREDENTIALS_FILE}" ]]; then
    if [[ "${CREDENTIALS_FILE}" == "-" ]]; then
      cat > "${tmp}"
      printf '%s\n' "${tmp}"
      return 0
    fi
    if [[ -f "${CREDENTIALS_FILE}" ]]; then
      printf '%s\n' "${CREDENTIALS_FILE}"
      return 0
    fi
    if [[ "${CREDENTIALS_FILE}" == curl\ * || "${CREDENTIALS_FILE}" == "{"* || "${CREDENTIALS_FILE}" == "["* ]]; then
      printf '%s\n' "${CREDENTIALS_FILE}" > "${tmp}"
      printf '%s\n' "${tmp}"
      return 0
    fi
    printf '%s\n' "${CREDENTIALS_FILE}"
    return 0
  fi

  if [[ "${PASTE_CREDENTIALS}" == "1" ]]; then
    read_pasted_credentials "${tmp}"
    printf '%s\n' "${tmp}"
    return 0
  fi

  cat >&2 <<EOF

Provide one authenticated widget request.

Accepted inputs:
  - Path to a Chrome DevTools "Copy as cURL" text file.
  - Path to a Chrome DevTools HAR export.
  - A one-line copied cURL command.
  - Type PASTE to paste a multi-line cURL/HAR block directly here.
  - Type - to read the request from stdin.

Good requests contain:
  - URL host: content-discoveryengine.googleapis.com
  - Header: Authorization: Bearer <widget JWT>
  - JSON body with configId: ${GE_WIDGET_CONFIG_ID}

Example path: /tmp/ge-widget-request.curl
EOF
  local input
  read -r -p "Path, one-line cURL, PASTE, or -: " input
  case "${input}" in
    PASTE|paste)
      read_pasted_credentials "${tmp}"
      printf '%s\n' "${tmp}"
      ;;
    -)
      cat > "${tmp}"
      printf '%s\n' "${tmp}"
      ;;
    curl\ *)
      write_curl_or_inline_request "${tmp}" "${input}"
      printf '%s\n' "${tmp}"
      ;;
    "{"*|"["*)
      printf '%s\n' "${input}" > "${tmp}"
      printf '%s\n' "${tmp}"
      ;;
    *)
      printf '%s\n' "${input}"
      ;;
  esac
}

read_pasted_credentials() {
  local output_file="$1"
  local sentinel="__GE_WIDGET_CURL_END__"
  : > "${output_file}"
  cat >&2 <<EOF

Paste the full cURL/HAR now.
Finish with a line containing only:
${sentinel}

EOF
  local line
  while IFS= read -r line; do
    if [[ "${line}" == "${sentinel}" ]]; then
      break
    fi
    printf '%s\n' "${line}" >> "${output_file}"
  done
}

write_curl_or_inline_request() {
  local output_file="$1"
  local first_line="$2"
  local line="${first_line}"
  printf '%s\n' "${line}" > "${output_file}"

  # If a multi-line cURL was pasted directly at the single-line prompt, subsequent lines are often
  # already queued on stdin. Capture continuation lines while the previous line ends in a backslash.
  while [[ "${line}" == *\\ ]]; do
    if ! IFS= read -r -t 0.25 line; then
      break
    fi
    printf '%s\n' "${line}" >> "${output_file}"
  done
}

extract_credentials() {
  section "Extract widget credentials"
  local input_file
  input_file="$(credentials_file)"
  if [[ ! -f "${input_file}" ]]; then
    echo "Credential file not found: ${input_file}" >&2
    exit 1
  fi

  python3 skill/extract_widget_credentials.py "${input_file}" \
    --env-file "${ENV_FILE}" \
    --token-file "${TOKEN_FILE}" \
    --project "${GE_PROJECT}" \
    --agent-planner "${GE_COMMAND_PLANNER_AGENT_ID}" \
    --agent-surface "${GE_SURFACE_COMMANDER_AGENT_ID}"

  # shellcheck disable=SC1090
  source "${ENV_FILE}"
  export GE_WIDGET_BEARER_TOKEN_FILE="${GE_WIDGET_BEARER_TOKEN_FILE:-${TOKEN_FILE}}"
}

list_skills() {
  section "List current private widget skills"
  cd "${ROOT_DIR}"
  local write_env_args=()
  if [[ "${WRITE_WEB_SHELL_ENV}" == "1" ]]; then
    write_env_args=(--write-env "${WEB_SHELL_ENV}")
  fi
  python3 skill/update_skills.py --api-mode widget --live --list "${write_env_args[@]}"
}

replace_skills() {
  cd "${ROOT_DIR}"
  if [[ "${DRY_RUN}" == "1" ]]; then
    section "Dry-run replacement plan"
    python3 skill/update_skills.py --api-mode widget --replace --yes
    return 0
  fi

  if [[ "${YES}" != "1" ]]; then
    cat <<EOF

This will DELETE and RECREATE visible widget skills matching:
  planner label:   m365-command-planner
  commander label: m365-surface-commander

Current GE_*_AGENT_ID values are treated as hints only:
  planner hint:   ${GE_COMMAND_PLANNER_AGENT_ID}
  commander hint: ${GE_SURFACE_COMMANDER_AGENT_ID}

The updater resolves the current visible numeric ids immediately before deleting. This avoids
stale-id failures after prior delete/recreate runs.

Type REPLACE to continue.
EOF
    local confirmation
    read -r -p "> " confirmation
    if [[ "${confirmation}" != "REPLACE" ]]; then
      echo "Aborted."
      exit 1
    fi
  fi

  section "Delete, recreate, upload, and verify latest skill bundles"
  local write_env_args=()
  if [[ "${WRITE_WEB_SHELL_ENV}" == "1" ]]; then
    write_env_args=(--write-env "${WEB_SHELL_ENV}")
  fi
  python3 skill/update_skills.py --api-mode widget --replace --yes --live "${write_env_args[@]}"

  section "List skills after replacement"
  python3 skill/update_skills.py --api-mode widget --live --list
}

main() {
  section "Target"
  cat <<EOF
project:        ${GE_PROJECT}
project_number: ${GE_PROJECT_NUMBER}
location:       ${GE_LOCATION}
engine:         ${GE_ENGINE}
widget_config:  ${GE_WIDGET_CONFIG_ID}
planner_id:     ${GE_COMMAND_PLANNER_AGENT_ID}
commander_id:   ${GE_SURFACE_COMMANDER_AGENT_ID}
env_file:       ${ENV_FILE}
token_file:     ${TOKEN_FILE}
min_token_ttl:  ${TOKEN_MIN_TTL_SECONDS}s
web_shell_env:  ${WEB_SHELL_ENV}
EOF

  rebuild_and_validate
  report_bundle_drift
  if ! reuse_widget_token_if_valid; then
    open_app_url
    extract_credentials
  fi
  list_skills
  if [[ "${LIST_ONLY}" == "1" ]]; then
    section "Done"
    if [[ "${WRITE_WEB_SHELL_ENV}" == "1" ]]; then
      echo "Listed current widget skills and synced ${WEB_SHELL_ENV}; no replacement performed."
    else
      echo "Listed current widget skills; no replacement performed."
    fi
    return 0
  fi
  replace_skills

  section "Next"
  if [[ "${WRITE_WEB_SHELL_ENV}" == "1" ]]; then
    echo "Updated ${WEB_SHELL_ENV}. Restart the dev server so Vite reloads the new skill refs."
  else
    cat <<'EOF'
Copy the printed VITE_GE_COMMAND_PLANNER_SKILL and VITE_GE_SURFACE_COMMANDER_SKILL
values into packages/web-shell/.env, then restart the dev server.
EOF
  fi
}

main "$@"
