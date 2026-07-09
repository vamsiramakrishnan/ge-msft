#!/usr/bin/env bash
# Wrapper for scripts/m365-tenant-addin.ps1.
#
# This is intentionally separate from the dev tunnel flow. Tenant deployment should usually use a
# stable origin/package, not a throwaway quick-tunnel URL.

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

usage() {
  cat <<'EOF'
Usage:
  scripts/m365-tenant-addin.sh <list|deploy|update|delete|assign|enable|disable> [ps1 args...]

Examples:
  scripts/m365-tenant-addin.sh list
  M365_ADDIN_MEMBERS='vamramak@psott.onmicrosoft.com' scripts/m365-tenant-addin.sh deploy
  M365_ADDIN_ASSIGN_EVERYONE=1 scripts/m365-tenant-addin.sh update
  scripts/m365-tenant-addin.sh deploy --dry-run

Environment:
  M365_ADDIN_PROFILE=development       Package profile to deploy.
  M365_ADDIN_MEMBERS=a@contoso.com,b@contoso.com
  M365_ADDIN_ASSIGN_EVERYONE=1         Explicitly assign the add-ins to everyone.
  M365_ADDIN_UPLOAD_ONLY=1             Upload manifests without assignments.
  M365_ADDIN_INSTALL_MODULE=1          Install PowerShell backend modules if missing.
  M365_ADDIN_SKIP_PACKAGE=1            Do not regenerate development manifests/packages first.
  M365_ADDIN_MANIFEST_PATH=...         Deploy/update/delete one manifest.
  M365_ADDIN_MANIFEST_DIR=...          Deploy/update/delete all manifests in a directory.
  M365_ADDIN_PRODUCT_ID=...            ProductId for update/delete/assign/enable/disable.
  M365_ADDIN_UPN=admin@tenant.com      Admin UPN for Connect-ExchangeOnline.
  M365_ADDIN_DEVICE=1                  Use device-code auth.
  M365_ADDIN_BACKEND=ExchangeOnlineManagement|O365CentralizedAddInDeployment|Auto

Notes:
  - Uses backend auto-detection by default and selects the module exposing New-OrganizationAddIn.
  - Targets add-in-only XML manifests in dist/package/<profile>/xml by default.
  - Unified Microsoft 365 manifests still require the Integrated Apps portal.
EOF
}

if [[ $# -lt 1 || "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

raw_action="$1"
shift

case "${raw_action}" in
  list) action="List" ;;
  deploy) action="Deploy" ;;
  update) action="Update" ;;
  delete) action="Delete" ;;
  assign) action="Assign" ;;
  enable) action="Enable" ;;
  disable) action="Disable" ;;
  *)
    echo "Unknown action: ${raw_action}" >&2
    usage >&2
    exit 2
    ;;
esac

if command -v pwsh >/dev/null 2>&1; then
  POWERSHELL_BIN="$(command -v pwsh)"
elif command -v powershell >/dev/null 2>&1; then
  POWERSHELL_BIN="$(command -v powershell)"
elif [[ "${M365_ADDIN_INSTALL_MODULE:-0}" == "1" && -x "${ROOT_DIR}/scripts/setup-pwsh-ubuntu.sh" ]]; then
  "${ROOT_DIR}/scripts/setup-pwsh-ubuntu.sh"
  if command -v pwsh >/dev/null 2>&1; then
    POWERSHELL_BIN="$(command -v pwsh)"
  else
    echo "PowerShell setup completed but pwsh is still not on PATH." >&2
    exit 1
  fi
else
  cat >&2 <<'EOF'
PowerShell is not installed on this machine.

Options:
  1. Install PowerShell Core (`pwsh`) here, then rerun this script.
  2. Run scripts/m365-tenant-addin.ps1 from Windows PowerShell as your Microsoft 365 admin.

The Microsoft module commands are:
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
  Install-Module -Name O365CentralizedAddInDeployment -Scope CurrentUser
EOF
  exit 1
fi

profile="${M365_ADDIN_PROFILE:-development}"

if [[ "${M365_ADDIN_SKIP_PACKAGE:-0}" != "1" ]]; then
  (
    cd "${ROOT_DIR}"
    bun run manifests:generate -- --profile "${profile}"
    bun run manifests:validate -- --profile "${profile}"
    case "${profile}" in
      development)
        bun run package:dev
        ;;
      alpha)
        bun run package:alpha
        ;;
      *)
        echo "No package script is known for profile '${profile}'. Set M365_ADDIN_SKIP_PACKAGE=1 if the package already exists." >&2
        exit 1
        ;;
    esac
  )
fi

ps_args=(-NoProfile -ExecutionPolicy Bypass -File "${ROOT_DIR}/scripts/m365-tenant-addin.ps1" -Action "${action}" -Profile "${profile}" -Backend "${M365_ADDIN_BACKEND:-Auto}")

if [[ -n "${M365_ADDIN_MANIFEST_PATH:-}" ]]; then
  ps_args+=(-ManifestPath "${M365_ADDIN_MANIFEST_PATH}")
fi
if [[ -n "${M365_ADDIN_MANIFEST_DIR:-}" ]]; then
  ps_args+=(-ManifestDir "${M365_ADDIN_MANIFEST_DIR}")
fi
if [[ -n "${M365_ADDIN_PRODUCT_ID:-}" ]]; then
  ps_args+=(-ProductId "${M365_ADDIN_PRODUCT_ID}")
fi
if [[ -n "${M365_ADDIN_LOCALE:-}" ]]; then
  ps_args+=(-Locale "${M365_ADDIN_LOCALE}")
fi
if [[ "${M365_ADDIN_ASSIGN_EVERYONE:-0}" == "1" ]]; then
  ps_args+=(-AssignToEveryone)
fi
if [[ "${M365_ADDIN_UPLOAD_ONLY:-0}" == "1" ]]; then
  ps_args+=(-UploadOnly)
fi
if [[ "${M365_ADDIN_INSTALL_MODULE:-0}" == "1" ]]; then
  ps_args+=(-InstallModule)
fi
if [[ "${M365_ADDIN_DEVICE:-0}" == "1" ]]; then
  ps_args+=(-Device)
fi
if [[ -n "${M365_ADDIN_UPN:-}" ]]; then
  ps_args+=(-UserPrincipalName "${M365_ADDIN_UPN}")
fi
if [[ -n "${M365_ADDIN_MEMBERS:-}" ]]; then
  IFS=',' read -r -a members <<< "${M365_ADDIN_MEMBERS}"
  ps_args+=(-Members)
  for member in "${members[@]}"; do
    trimmed="$(printf '%s' "${member}" | xargs)"
    if [[ -n "${trimmed}" ]]; then
      ps_args+=("${trimmed}")
    fi
  done
fi

while [[ $# -gt 0 ]]; do
  case "$1" in
    --dry-run)
      ps_args+=(-DryRun)
      shift
      ;;
    --install-module)
      ps_args+=(-InstallModule)
      shift
      ;;
    --upload-only)
      ps_args+=(-UploadOnly)
      shift
      ;;
    --assign-everyone)
      ps_args+=(-AssignToEveryone)
      shift
      ;;
    --no-connect)
      ps_args+=(-NoConnect)
      shift
      ;;
    --device)
      ps_args+=(-Device)
      shift
      ;;
    --backend)
      ps_args+=(-Backend "${2:-}")
      shift 2
      ;;
    --user-principal-name)
      ps_args+=(-UserPrincipalName "${2:-}")
      shift 2
      ;;
    *)
      ps_args+=("$1")
      shift
      ;;
  esac
done

"${POWERSHELL_BIN}" "${ps_args[@]}"
