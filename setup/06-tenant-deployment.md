# Tenant Deployment Automation

Use this flow when you want Microsoft 365 / Office on the web users to receive the add-in through
tenant deployment instead of manually uploading the manifest in each browser.

## What This Automates

The repo includes a wrapper around the Microsoft 365 Centralized Deployment add-in cmdlets:

```bash
scripts/m365-tenant-addin.sh
scripts/m365-tenant-addin.ps1
```

By default, the wrapper targets the Claude-style add-in-only XML pair generated into:

```text
dist/package/<profile>/centralized/
```

For development this is usually:

```text
dist/package/development/centralized/office.manifest.xml
dist/package/development/centralized/outlook.manifest.xml
```

`office.manifest.xml` uses one stable add-in identity for Word, Excel, and PowerPoint and declares
host-specific Home-ribbon commands. Outlook remains separate because Microsoft uses the `MailApp`
schema for it. The host-specific manifests under `dist/package/development/xml/` remain available
for manual/web diagnostics but aren't deployed by default, which avoids duplicate ribbon entries.

## Requirements

- Microsoft 365 tenant that supports Centralized Deployment.
- An admin account with the required Microsoft 365 / Exchange admin permissions.
- PowerShell Core (`pwsh`) on Linux, or PowerShell on Windows.
- The Microsoft 365 PowerShell modules used by backend auto-detection:

```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
Install-Module -Name O365CentralizedAddInDeployment -Scope CurrentUser
```

On Ubuntu, PowerShell Core can be installed with:

```bash
mise run m365:pwsh:setup
```

Leave `-Backend Auto` unless you are debugging. The script imports available modules and selects the
one that exposes `New-OrganizationAddIn`. Some `ExchangeOnlineManagement` installs connect
successfully but do not expose the centralized add-in deployment cmdlets, so module presence alone is
not enough.

On Linux, there is one known bad split:

- `ExchangeOnlineManagement` may support device-code login but not expose `New-OrganizationAddIn`.
- `O365CentralizedAddInDeployment` may expose `New-OrganizationAddIn` but fail during
  `Connect-OrganizationAddInService` with a `kernel32.dll` load error.

When that happens, Linux can still generate packages and run `-NoConnect -DryRun`, but live tenant
deployment needs PowerShell on Windows or the Microsoft 365 admin center Integrated Apps /
Centralized Deployment portal.

`bun bootstrap` treats this as a guarded handoff. It still generates packages and updates Gemini
Enterprise skills, but when the local Linux runtime cannot complete live tenant deployment it skips
that phase, prints the Windows PowerShell command to run, and does not mark the manifest SHA as
tenant-deployed in `.ge-deploy.json`.

Use `--force-tenant-deploy` only when you are intentionally testing a new Microsoft module/runtime
combination. Use `--skip-tenant-deploy` when you only want packaging and skill provisioning.

To explicitly test the cross-platform Exchange backend:

```bash
bun bootstrap -- --tenant-backend ExchangeOnlineManagement --force-tenant-deploy
```

That only works when this check returns `New-OrganizationAddIn`:

```bash
pwsh -NoProfile -Command "Import-Module ExchangeOnlineManagement -Force; Get-Command New-OrganizationAddIn"
```

If it prints nothing, the Exchange backend can authenticate but cannot deploy add-ins in that
session.

## List Existing Add-ins

From Linux:

```bash
M365_ADDIN_UPN='admin@psott.onmicrosoft.com' M365_ADDIN_DEVICE=1 mise run m365:addins:list
```

Equivalent direct PowerShell:

```powershell
./scripts/m365-tenant-addin.ps1 -Action List -Backend Auto -UserPrincipalName 'admin@psott.onmicrosoft.com' -Device
```

## Deploy to Specific Users or Groups

Start with yourself or a test group:

```powershell
./scripts/m365-tenant-addin.ps1 `
  -Action Deploy `
  -Backend Auto `
  -InstallModule `
  -UserPrincipalName 'admin@psott.onmicrosoft.com' `
  -Device `
  -Members 'vamramak@psott.onmicrosoft.com'
```

From `mise`:

```bash
M365_ADDIN_UPN='admin@psott.onmicrosoft.com' \
M365_ADDIN_DEVICE=1 \
M365_ADDIN_INSTALL_MODULE=1 \
M365_ADDIN_MEMBERS='vamramak@psott.onmicrosoft.com' \
mise run m365:addins:deploy
```

Multiple users/groups are comma-separated:

```powershell
./scripts/m365-tenant-addin.ps1 `
  -Action Deploy `
  -Backend Auto `
  -InstallModule `
  -UserPrincipalName 'admin@psott.onmicrosoft.com' `
  -Device `
  -Members 'vamramak@psott.onmicrosoft.com','ge-testers@psott.onmicrosoft.com'
```

The script refuses to deploy without one of:

- `M365_ADDIN_MEMBERS`
- `M365_ADDIN_ASSIGN_EVERYONE=1`
- `M365_ADDIN_UPLOAD_ONLY=1`

This prevents accidentally creating an unclear tenant rollout.

## Deploy to Everyone

Only use this when the manifest points at a stable origin and the add-in has been tested:

```powershell
./scripts/m365-tenant-addin.ps1 `
  -Action Deploy `
  -Backend Auto `
  -InstallModule `
  -UserPrincipalName 'admin@psott.onmicrosoft.com' `
  -Device `
  -AssignToEveryone
```

## Update Existing Deployed Add-ins

When the manifest changed, regenerate the package and update:

```powershell
./scripts/m365-tenant-addin.ps1 `
  -Action Update `
  -Backend Auto `
  -UserPrincipalName 'admin@psott.onmicrosoft.com' `
  -Device `
  -Members 'vamramak@psott.onmicrosoft.com'
```

If the web code changed but the manifest URL stayed the same, tenant redeployment is usually not
needed. Users load the web app from the manifest URL.

## Delete Deployed Add-ins

```bash
mise run m365:addins:delete
```

By default, the script derives product IDs from the generated XML manifest `<Id>` values. To delete a
single add-in by product ID:

```powershell
./scripts/m365-tenant-addin.ps1 `
  -Action Delete `
  -Backend Auto `
  -UserPrincipalName 'admin@psott.onmicrosoft.com' `
  -Device `
  -ProductId '<product-guid>'
```

## Dry Run

```powershell
./scripts/m365-tenant-addin.ps1 `
  -Action Deploy `
  -Backend Auto `
  -Members 'vamramak@psott.onmicrosoft.com' `
  -NoConnect `
  -DryRun
```

## Stable Origin Requirement

Do not use a Cloudflare quick-tunnel origin for broad tenant deployment. Quick tunnel hostnames are
ephemeral, and every hostname change requires:

- Entra redirect update.
- Manifest regeneration.
- Tenant add-in update.
- User/client propagation.

Use quick tunnels for dev sideloading. Use a stable HTTPS origin for tenant rollout.
See [Hosting origin and release flow](./08-hosting-origin-and-release.md) for the exact Office load
model and the tradeoffs between GCS/CDN, App Engine, and Cloud Run.

The development central lane also requires stable, distinct public GUIDs:

```bash
GE_DEV_OFFICE_XML_APP_ID='<stable-guid-for-word-excel-powerpoint>'
GE_DEV_OUTLOOK_APP_ID='<stable-guid-for-outlook>'
```

## Entra SPA Redirect Sync

Every manifest/package profile has an auth redirect:

```text
https://<origin>/auth-redirect.html
```

For development, the origin comes from `GE_DEV_WEB_ORIGIN` / `GE_DEV_WEB_DOMAIN` in
`packages/web-shell/.env`. For the release profile, the origin comes from `GE_ALPHA_WEB_DOMAIN`.

Sync redirects explicitly:

```bash
bun run entra:sync:dev
bun run entra:sync:release
```

The sync tool:

- finds the Entra application by the profile's client ID,
- adds the profile redirect URI to `spa.redirectUris`,
- removes stale `*.trycloudflare.com/auth-redirect.html` redirects unless
  `--keep-stale-redirects` is passed,
- writes the planned patch to `.ge-dev/entra-spa-patch.<profile>.json`.

`bun run bootstrap:release` runs this before catalog upload. If it fails, do not upload the package;
the add-in may install but MSAL redirect handling will break.

## Unified Manifest Note

This script uses Centralized Deployment cmdlets for add-in-only XML manifests. For the unified
Microsoft 365 package zip, use the separate app-catalog lane:

```bash
bun run m365:catalog
```

See [Deployment methods matrix](./07-deployment-methods-matrix.md) before choosing between XML
Centralized Deployment and unified package catalog upload.
