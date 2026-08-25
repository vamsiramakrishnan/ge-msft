#!/usr/bin/env bash
# Upload the unified Microsoft 365 app package through CLI for Microsoft 365.
#
# This targets the Teams/M365 app catalog lane (`dist/release/*-m365-v*.zip`), not the classic
# Office XML Centralized Deployment lane.

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

usage() {
  cat <<'EOF'
Usage:
  scripts/m365-app-catalog.sh <status|login|list|add|update|upsert|remove> [--dry-run]

Environment:
  M365_APP_PACKAGE=dist/release/development-m365-v<version>.zip
  M365_APP_PROFILE=development       Package profile when M365_APP_PACKAGE is not set.
  M365_CLI_INSTALL=1              Install @pnp/cli-microsoft365 with Bun if m365 is missing.
  M365_CLI_LOGIN=1                Run m365 login before the action.
  M365_CLI_AUTH_TYPE=deviceCode   Login auth type for m365 login.
  M365_APP_ID=<guid>              Override app id for update/remove/upsert.

Notes:
  - add/update/upsert use the unified package zip, not XML manifests.
  - upsert checks m365 teams app get --id <manifest id>, then update or add.
EOF
}

if [[ $# -lt 1 || "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

action="$1"
shift
dry_run=0
while [[ $# -gt 0 ]]; do
  case "$1" in
    --dry-run)
      dry_run=1
      shift
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

run() {
  if [[ "${dry_run}" == "1" ]]; then
    printf '[dry-run]'
    printf ' %q' "$@"
    printf '\n'
    return
  fi
  "$@"
}

resolve_package() {
  if [[ -n "${M365_APP_PACKAGE:-}" ]]; then
    printf '%s\n' "${M365_APP_PACKAGE}"
    return
  fi
  M365_APP_PROFILE="${M365_APP_PROFILE:-development}" node - <<'NODE'
const fs = require('fs');
const path = require('path');
const root = process.cwd();
const profile = process.env.M365_APP_PROFILE || 'development';
if (profile === 'development') {
  const artifact = path.join(root, 'dist/release/development-artifact.json');
  if (fs.existsSync(artifact)) {
    const parsed = JSON.parse(fs.readFileSync(artifact, 'utf8'));
    if (parsed.m365Package) {
      console.log(parsed.m365Package);
      process.exit(0);
    }
  }
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  console.log(path.join(root, 'dist/release', `development-m365-v${pkg.version}.zip`));
  process.exit(0);
}
const artifact = path.join(root, 'dist/release/artifact.json');
if (fs.existsSync(artifact)) {
  const parsed = JSON.parse(fs.readFileSync(artifact, 'utf8'));
  if (parsed.profile === profile && parsed.package) {
    console.log(parsed.package);
    process.exit(0);
  }
}
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
console.log(path.join(root, 'dist/release', `${profile}-v${pkg.version}.zip`));
NODE
}

manifest_id() {
  local package_path="$1"
  if [[ -n "${M365_APP_ID:-}" ]]; then
    printf '%s\n' "${M365_APP_ID}"
    return
  fi
  unzip -p "${package_path}" manifest.json | node -e 'let text = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", chunk => { text += chunk; });
process.stdin.on("end", () => {
  const parsed = JSON.parse(text);
  if (!parsed.id) process.exit(1);
  console.log(parsed.id);
});'
}

assert_catalog_ready() {
  local package_path="$1"
  local app_id="$2"
  local manifest_json
  manifest_json="$(unzip -p "${package_path}" manifest.json)"
  local valid_domains
  valid_domains="$(printf '%s' "${manifest_json}" | node -e 'let text = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", chunk => { text += chunk; });
process.stdin.on("end", () => {
  const m = JSON.parse(text);
  console.log((m.validDomains || []).join(","));
});')"
  if [[ "${M365_ALLOW_DEV_PLACEHOLDERS:-0}" != "1" && "${app_id}" =~ ^11111111-1111-4111-8111-111111111111$ ]]; then
    cat >&2 <<EOF
Refusing live catalog upload with the development placeholder app id:
  ${app_id}

Set GE_DEV_APP_ID to a stable tenant-owned GUID, rerun bun run setup:package, then retry.
For throwaway dev-only testing, set M365_ALLOW_DEV_PLACEHOLDERS=1.
EOF
    exit 1
  fi
  if [[ "${M365_ALLOW_QUICK_TUNNEL:-0}" != "1" && "${valid_domains}" == *".trycloudflare.com"* ]]; then
    cat >&2 <<EOF
Refusing live catalog upload with a Cloudflare quick-tunnel domain:
  ${valid_domains}

Use a stable HTTPS origin for catalog deployment, rerun bun run setup:package, then retry.
For throwaway dev-only testing, set M365_ALLOW_QUICK_TUNNEL=1.
EOF
    exit 1
  fi
}

if ! command -v m365 >/dev/null 2>&1; then
  if [[ "${M365_CLI_INSTALL:-0}" != "1" ]]; then
    if [[ "${dry_run}" == "1" ]]; then
      echo "[dry-run] m365 is not installed; command construction will still be printed"
    else
    cat >&2 <<'EOF'
CLI for Microsoft 365 (`m365`) is not installed.

Install it with:
  bun add -g @pnp/cli-microsoft365

Or rerun with:
  M365_CLI_INSTALL=1 scripts/m365-app-catalog.sh <action>
EOF
    exit 1
    fi
  fi
  if [[ "${M365_CLI_INSTALL:-0}" == "1" ]]; then
    run bun add -g @pnp/cli-microsoft365
  fi
fi

package_path="$(resolve_package)"
if [[ "${action}" != "status" && "${action}" != "login" && "${action}" != "list" && ! -f "${package_path}" ]]; then
  echo "Unified package not found: ${package_path}" >&2
  echo "Run bun run setup:package for development, or bun run bootstrap:release:dry-run to validate release profile packaging." >&2
  exit 1
fi

if [[ "${M365_CLI_LOGIN:-0}" == "1" && "${action}" != "login" ]]; then
  run m365 login --authType "${M365_CLI_AUTH_TYPE:-deviceCode}"
fi

case "${action}" in
  status)
    run m365 status
    ;;
  login)
    run m365 login --authType "${M365_CLI_AUTH_TYPE:-deviceCode}"
    ;;
  list)
    run m365 teams app list
    ;;
  add)
    assert_catalog_ready "${package_path}" "$(manifest_id "${package_path}")"
    run m365 teams app add --filePath "${package_path}"
    ;;
  update)
    app_id="$(manifest_id "${package_path}")"
    assert_catalog_ready "${package_path}" "${app_id}"
    run m365 teams app update --id "${app_id}" --filePath "${package_path}"
    ;;
  upsert)
    app_id="$(manifest_id "${package_path}")"
    if [[ "${dry_run}" != "1" ]]; then
      assert_catalog_ready "${package_path}" "${app_id}"
    fi
    if [[ "${dry_run}" == "1" ]]; then
      run m365 teams app get --id "${app_id}"
      run m365 teams app update --id "${app_id}" --filePath "${package_path}"
    elif m365 teams app get --id "${app_id}" >/dev/null 2>&1; then
      run m365 teams app update --id "${app_id}" --filePath "${package_path}"
    else
      run m365 teams app add --filePath "${package_path}"
    fi
    ;;
  remove)
    app_id="$(manifest_id "${package_path}")"
    run m365 teams app remove --id "${app_id}" --confirm
    ;;
  *)
    echo "Unknown action: ${action}" >&2
    usage >&2
    exit 2
    ;;
esac
