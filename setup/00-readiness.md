# Readiness and Guided Setup

Use the guided setup CLI before changing manifests, tunnels, tenant deployment, or Gemini
Enterprise skills.

## First Command

```bash
bun install
bun run setup:guide
```

This opens an interactive Bun-powered CLI for:

- Readiness checks.
- Azure / Entra login checks.
- Exchange Online tenant add-in deployment checks.
- Google WIF login checks.
- Gemini Enterprise widget credential flow.
- Manifest generation and packaging.
- Vite + Cloudflare tunnel + Entra redirect registration.
- Gemini Enterprise skill list/update flows.

## Install Or Repair Prerequisites

Use this when a workstation is new or the doctor reports missing tools:

```bash
bun run setup:prereqs
```

Target one area when you do not want the full guided pass:

```bash
bun run setup:prereqs -- --target deps
bun run setup:prereqs -- --target azure
bun run setup:prereqs -- --target gcloud
bun run setup:prereqs -- --target atk
bun run setup:prereqs -- --target powershell,exchange
bun run setup:prereqs -- --target cloudflared
bun run setup:prereqs -- --target package
```

Short aliases are available for the common cases:

```bash
bun run setup:deps
bun run setup:azure
bun run setup:gcloud
bun run setup:gcloud:wif
bun run setup:atk
bun run setup:powershell
bun run setup:cloudflared
```

Bun orchestrates the workflow. It still delegates system packages to their native installer:
Azure CLI installs into the repo-local `.venv-az` with a `bin/az` wrapper, Google Cloud CLI uses
the official Ubuntu apt repository, Microsoft 365 Agents Toolkit installs with `bun add -g`,
PowerShell on Ubuntu uses `sudo apt-get`, Microsoft 365 add-in deployment modules use
`pwsh Install-Module`, and Cloudflare Tunnel downloads the `cloudflared` binary to
`/tmp/cloudflared`.

## Google Auth Split

Keep two Google auth lanes separate:

- **Global ADC / Claude Code / Vertex work:** do not create a repo WIF ADC file. On this
  workstation, leave `~/.config/gcloud/application_default_credentials.json` absent so Google auth
  libraries fall through to the VM metadata default service account for `vital-octagon-19612`.
- **Gemini Enterprise Saibalaji WIF:** use the repo-local gcloud config at `.gcloud/`. This signs in
  to the Entra-backed workforce pool for `saib-ai-playground` without touching ADC.

Run the WIF login flow with:

```bash
bun run setup:gcloud:wif
# or
mise run setup:gcloud:wif
```

The command first checks whether the repo-local gcloud config can already mint an access token. If
that works, it skips browser login. If the cached WIF session is missing or expired, it runs
`gcloud auth login --login-config ... --no-launch-browser` with `CLOUDSDK_CONFIG=$PWD/.gcloud`,
sets the repo-local project/quota project, and intentionally never runs
`gcloud auth application-default login`.

Force a fresh browser login only when switching accounts or repairing local auth state:

```bash
bun run setup:gcloud:wif:force
# or
mise run setup:gcloud:wif:force
```

Token lifetime is controlled by Entra, the workforce identity provider, and Google Cloud. The setup
CLI does not try to extend tenant policy; it reuses and refreshes what the identity providers allow,
then prompts only when refresh fails.

## Non-interactive Doctor

```bash
bun run setup:doctor
```

JSON output for automation:

```bash
bun run setup:doctor -- --json
```

Offer guided fixes:

```bash
bun run setup:doctor:fix
```

## Package Manifests

```bash
bun run setup:package
```

## One-Shot Bootstrap

Use bootstrap for the full local operator flow:

```bash
bun bootstrap
```

Use the named lanes for day-to-day work:

```bash
bun run bootstrap:dev             # safe dev bootstrap: package + skills, no tenant/catalog deploy
bun run bootstrap:dev:sideload    # package and install the unified dev package with atk
bun run bootstrap:release:dry-run # inspect the release path without mutating tenant/catalog state
bun run bootstrap:release         # stable-origin release profile + unified catalog upsert
bun run bootstrap:prod            # alias for bootstrap:release
```

`bootstrap:prod` deliberately does **not** mean "upload the development manifest everywhere." It is
an alias for the current production-like release profile, `internal-alpha-word-excel`, which requires
stable `GE_ALPHA_*` values and blocks placeholder/localhost release config during manifest
generation.

The release path is wired in this order:

```text
GE_ALPHA_* release config
  -> build web shell
  -> generate + validate release manifest
  -> package release zip
  -> sync Entra SPA redirect: https://GE_ALPHA_WEB_DOMAIN/auth-redirect.html
  -> upload/upsert the profile-specific zip through the selected Microsoft lane
```

Direct Entra redirect sync commands are also available:

```bash
bun run entra:sync:dev
bun run entra:sync:release
```

Use `--skip-entra` on `bun bootstrap` only when the redirect is already known to be correct.

By default it uses `--deployment-lane auto`: XML Centralized Deployment when the local Microsoft
PowerShell backend can actually run it, otherwise unified Microsoft 365 app catalog upload. On this
Linux workstation, the automated deploy lane is the catalog path:

```bash
bun bootstrap -- --deployment-lane catalog
```

Use XML explicitly only on an environment where `New-OrganizationAddIn` is available and can
authenticate:

```bash
bun bootstrap -- --deployment-lane xml --tenant-backend Auto
```

This runs:

```bash
bun run manifests:generate -- --profile development
bun run manifests:validate -- --profile development
bun run package:dev
```

## Dev Tunnel and Entra Redirect

Guided:

```bash
bun run ge:dev:tunnel
```

Mise wrapper:

```bash
mise run ge:dev:tunnel
```

The task restarts Vite, starts a Cloudflare quick tunnel, writes the new origin to
`packages/web-shell/.env`, regenerates manifests, and patches the Entra SPA redirect URI.

## Tenant Add-in Deployment

Before choosing a lane, read [Deployment methods matrix](./07-deployment-methods-matrix.md). It
separates Office add-ins, unified Microsoft 365 app packages, app catalog upload, sideloading, and
tenant-wide deployment.

Guided:

```bash
bun run m365:addins
```

Direct test-user deployment from Linux:

```bash
M365_ADDIN_UPN='admin@psott.onmicrosoft.com' \
M365_ADDIN_DEVICE=1 \
M365_ADDIN_INSTALL_MODULE=1 \
M365_ADDIN_MEMBERS='vamramak@psott.onmicrosoft.com' \
bun run m365:addins -- deploy --members 'vamramak@psott.onmicrosoft.com' --device --upn 'admin@psott.onmicrosoft.com'
```

Tenant-wide deployment:

```bash
M365_ADDIN_UPN='admin@psott.onmicrosoft.com' \
M365_ADDIN_DEVICE=1 \
M365_ADDIN_INSTALL_MODULE=1 \
M365_ADDIN_ASSIGN_EVERYONE=1 \
bun run m365:addins -- deploy --assignment everyone --device --upn 'admin@psott.onmicrosoft.com'
```

This uses backend auto-detection. The scripts install/check both `ExchangeOnlineManagement` and
`O365CentralizedAddInDeployment`, then select the module that actually exposes
`New-OrganizationAddIn` in the current PowerShell session.

For the unified Microsoft 365 package app catalog lane:

```bash
bun run m365:catalog
```

## Gemini Enterprise Skills

Guided:

```bash
bun run ge:skills
```

Direct flows:

```bash
bun run ge:skills -- --mode list
bun run ge:skills -- --mode dry-run
bun run ge:skills -- --mode update
```

Use `bun run ge:skills -- --mode paste-update` when you want to paste a full DevTools cURL/HAR block and let the
script extract the short-lived widget bearer token.

## Bun First

This repo uses Bun as the package manager and task runner:

```bash
bun install
bun run setup:prereqs
bun run setup:doctor
bun run setup:package
bun run m365:addins
bun run test
bun run build
```

The generated `bun.lock` is the package lock. Do not regenerate `package-lock.json`.
The `mise` tasks mirror the Bun commands for operators who prefer mise, but Bun is the primary
interface for local add-in development, packaging, tenant deployment guidance, and skill updates.
