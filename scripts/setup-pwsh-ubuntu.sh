#!/usr/bin/env bash
# Install PowerShell Core on Ubuntu and optionally install the Microsoft 365 PowerShell modules
# used by the tenant add-in deployment wrapper. Different tenants/module versions expose the
# centralized deployment cmdlets through different modules, so install both and let the wrapper
# select by cmdlet availability.

set -euo pipefail

if [[ -r /etc/os-release ]]; then
  # shellcheck disable=SC1091
  source /etc/os-release
else
  echo "/etc/os-release not found; this setup script only supports Ubuntu." >&2
  exit 1
fi

if [[ "${ID:-}" != "ubuntu" ]]; then
  echo "Unsupported OS: ${PRETTY_NAME:-unknown}. This setup script only supports Ubuntu." >&2
  exit 1
fi

if ! command -v pwsh >/dev/null 2>&1; then
  echo "Installing PowerShell Core for Ubuntu ${VERSION_ID}"
  sudo apt-get update
  sudo apt-get install -y wget apt-transport-https software-properties-common
  tmp_deb="$(mktemp -t packages-microsoft-prod.XXXXXX.deb)"
  wget -q "https://packages.microsoft.com/config/ubuntu/${VERSION_ID}/packages-microsoft-prod.deb" -O "${tmp_deb}"
  sudo dpkg -i "${tmp_deb}"
  rm -f "${tmp_deb}"
  sudo apt-get update
  sudo apt-get install -y powershell
else
  echo "PowerShell already installed: $(pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()')"
fi

if [[ "${M365_ADDIN_INSTALL_MODULE:-1}" == "1" ]]; then
  echo "Installing/updating Microsoft 365 add-in PowerShell modules for the current user"
  pwsh -NoProfile -Command "'ExchangeOnlineManagement','O365CentralizedAddInDeployment' | ForEach-Object { Install-Module -Name $_ -Scope CurrentUser -Force }"
fi

cat <<'EOF'

PowerShell setup complete.

Next:
  Use device-code auth from Linux when browser popups are awkward:

  bun run m365:addins -- list --upn 'admin@tenant.com' --device

Deploy to a test user:

  bun run m365:addins -- deploy --upn 'admin@tenant.com' --device \
    --assignment members --members 'user@tenant.com'

EOF
