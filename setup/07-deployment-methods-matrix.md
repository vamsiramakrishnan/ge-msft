# Deployment Methods Matrix

This repo produces multiple artifacts because Microsoft 365 has multiple deployment lanes. Do not
treat them as interchangeable.

## Terminology Model

| Term | What it means | Concrete artifact or place | Who controls it | What it does not mean |
| --- | --- | --- | --- | --- |
| Office Add-in | A web app surfaced inside an Office host such as Word, Excel, PowerPoint, Outlook, or OneNote through Office.js and a manifest | XML add-in-only manifest, or an Office extension inside a unified M365 manifest | Developer defines it; user/admin installs it | It is not automatically a Teams app; it is not native desktop code |
| Add-in-only manifest | The classic XML manifest format for Office Add-ins | `*.manifest.xml` | Developer/package pipeline | It is not the unified M365 app package zip |
| Unified app manifest for Microsoft 365 | JSON manifest that can package Office extensions and Teams/M365 app capabilities into one app package | `manifest.json` inside `development-m365-v<version>.zip` | Developer/package pipeline | It is not supported everywhere classic XML is supported |
| App for Microsoft 365 | A broader app package concept that can include Teams app surfaces and Office extensions | Unified package zip with `manifest.json` and icons | Developer/package pipeline; tenant admin governs availability | It is not necessarily installed into every Office ribbon just because it is uploaded |
| Teams app | An app package surface for Microsoft Teams, often also the packaging lane for broader M365 apps | Teams/M365 app package zip | Developer and Teams admin policies | It is not the same thing as a classic Office XML add-in |
| App catalog | A tenant catalog/repository where app packages are uploaded or made available | Teams/M365 app catalog, SharePoint app catalog, Microsoft 365 admin app inventory | Tenant admin | Catalog presence is not the same as assignment, pinning, or guaranteed install |
| Integrated Apps portal | Microsoft 365 admin center experience for deploying/managing apps and add-ins | Admin center UI | Tenant admin | It is not a local dev sideload path |
| Centralized Deployment | Admin deployment/assignment mechanism for Office Add-ins to users/groups/everyone | Admin center or PowerShell cmdlets such as `New-OrganizationAddIn` where available | Tenant admin / Exchange admin | It is not the Teams app catalog upload command |
| Sideloading | Developer/test install into a specific client/user/browser | Upload manifest dialog, dev tools, Agents Toolkit install | Developer/test user | It is not tenant rollout |
| Installation | The app/add-in is available to a user/client and can appear in Office UI | User/client state after sideload or admin deployment | User/admin/client | Uploading to a catalog alone may not install it |
| Assignment | Admin targets users/groups/everyone for an add-in/app | Centralized Deployment assignment or admin center assignment | Tenant admin | It is not just uploading a package |
| Pinning | Admin policy makes an app appear in a prominent host location, especially Teams | Teams app setup policies | Teams admin | It is not equivalent to Office ribbon command deployment |

## Mental Model

Think of the system as four layers:

| Layer | Question | Examples in this repo |
| --- | --- | --- |
| Artifact | What file did we build? | XML manifests, OneNote XML, unified M365 zip, web shell |
| Catalog/deployment lane | Where can that file be uploaded? | Office add-in upload dialog, Integrated Apps portal, Centralized Deployment PowerShell, Teams/M365 app catalog |
| Assignment/availability | Which users can see or install it? | Assigned user/group/everyone, Teams app policies, app catalog availability |
| Runtime/client | Where does it actually run? | Office on the web, Windows desktop Office, Mac desktop Office, Outlook, Teams, OneNote |
| Hosting origin | Where does Office download the add-in web app from? | Cloudflare dev tunnel, GCS/CDN, App Engine, Cloud Run, or another stable HTTPS host |

Most deployment mistakes happen when these layers are collapsed. For example:

- `m365 teams app add --filePath development-m365-v<version>.zip` uploads a unified app package to an
  app catalog lane. It does not deploy the XML manifests in `dist/package/development/xml`.
- `New-OrganizationAddIn -ManifestPath excel.manifest.xml` deploys a classic Office XML add-in. It
  does not upload the unified M365 package zip.
- A package being in a catalog does not prove every target Office client supports the manifest,
  command surface, runtime requirement set, or assignment policy.

## Add-in Versus App Catalog

| Question | Office Add-in | App catalog |
| --- | --- | --- |
| What is it? | The actual Office extension experience: task pane, ribbon commands, context menu commands, event handlers, custom functions, etc. | A tenant/package repository or admin inventory where app packages can be uploaded, approved, and made available |
| Is it a file? | No. It is represented by a manifest plus hosted web assets | No. It is a service/location that stores or references app packages |
| What file does it use here? | XML manifests or the Office extension part of the unified `manifest.json` | Unified M365 package zip for Teams/M365 catalog lanes; XML/add-in packages through admin center lanes |
| Does upload mean users get it? | Not by itself | Not by itself; admins may still need assignment, policies, consent, or user install |
| Does it define UI? | Yes, via manifest command/ribbon/context entries and Office.js runtime code | No, it stores/distributes packages that define UI |
| Does it define permissions? | The manifest can request Graph/resource permissions and Office requirement sets | The catalog/admin flow may approve, block, or govern the app; it is not the source of the add-in code |
| Is it host-specific? | Yes. Word/Excel/PPT/Outlook/OneNote/Teams have different support and APIs | Catalogs are broader, but clients still decide whether they can run a package |

In our repo:

| Repo thing | Concept |
| --- | --- |
| `packages/web-shell` | The hosted web app/runtime for the add-in |
| `dist/package/development/centralized/office.manifest.xml` | Centralized Word + Excel + PowerPoint Office Add-in manifest |
| `dist/package/development/centralized/outlook.manifest.xml` | Separate centralized Outlook Office Add-in manifest |
| `dist/package/development/xml/*.manifest.xml` | Host-specific XML manifests retained for manual diagnostics |
| `dist/package/development/onenote/onenote.manifest.xml` | Classic OneNote add-in manifest |
| `dist/release/development-m365-v<version>.zip` | Unified app package for Microsoft 365 / Teams-style app catalog lanes |
| `scripts/m365-tenant-addin.sh` | Classic XML Centralized Deployment automation |
| `scripts/m365-app-catalog.sh` | Unified package app catalog upload automation |

## Artifact Summary

| Artifact | Generated path | Primary use | Not for |
| --- | --- | --- | --- |
| Unified Microsoft 365 package | `dist/release/development-m365-v<version>.zip` | M365/Teams app package upload, Agents Toolkit, CLI for Microsoft 365 app catalog flows | Legacy Office XML upload dialogs |
| Centralized Office XML pair | `dist/package/development/centralized/{office,outlook}.manifest.xml` | Claude-style Centralized Deployment through Integrated Apps or PowerShell | Teams app catalog upload |
| Host-specific XML manifests | `dist/package/development/xml/*.manifest.xml` | Manual Office add-in sideloading and diagnostics | Default tenant-wide deployment |
| OneNote XML manifest | `dist/package/development/onenote/onenote.manifest.xml` | OneNote legacy add-in flow | Unified manifest flow |
| Web shell build | `dist/package/development/web/` | Hosted task pane/runtime assets referenced by manifests | Direct deployment by itself |

The web shell build is necessary but not sufficient. Office needs a manifest/package that points at
the hosted origin, and the hosted origin must serve the built web shell over HTTPS. See
[Hosting origin and release flow](./08-hosting-origin-and-release.md) for the App Engine, Cloud Run,
GCS/CDN, and Cloudflare tunnel decision.

## Current Repo State

As of the development package generated by `bun run package:dev`:

| Check | Current answer |
| --- | --- |
| Unified package built? | Yes: `dist/release/development-m365-v1.0.0.zip` |
| Unified manifest validates? | Yes: `bun run manifests:validate -- --profile development` passes |
| Unified package contents | `manifest.json` plus required icon assets |
| Unified manifest scopes | `mail`, `workbook`, `document`, `presentation` |
| Unified runtime URL | Current dev tunnel / configured web origin |
| Ribbon command | One `Open Gemini` button on the host home ribbon |
| OneNote in unified package? | No. OneNote uses the separate XML manifest |
| Production-ready catalog identity? | Not by default. Development uses `GE_DEV_APP_ID` or a placeholder GUID. Set a stable, tenant-owned GUID before serious catalog deployment |
| Production-ready origin? | Only if `GE_DEV_WEB_ORIGIN` / profile config points at a stable HTTPS origin. A Cloudflare quick tunnel is dev-only |

Plainly:

- **Yes, we build a unified add-in package.**
- **Yes, that package can be uploaded through an app catalog lane.**
- **It should work in supported Microsoft 365 clients for the scopes in the unified manifest.**
- **It does not cover OneNote.**
- **It does not replace XML Centralized Deployment for older/classic Office add-in coverage.**
- **For real tenant rollout, do not use placeholder IDs or ephemeral tunnel origins.**

Before a serious catalog upload, set stable development/release values:

```bash
GE_DEV_APP_ID='<stable-guid-for-this-m365-app-package>'
GE_DEV_WEB_ORIGIN='https://<stable-host>'
GE_DEV_WEB_DOMAIN='<stable-host>'
GE_DEV_ENTRA_CLIENT_ID='<entra-client-id>'
```

Then regenerate:

```bash
bun run setup:package
```

## Deployment Lanes

| Lane | Automation | Scope | Artifact | Linux-friendly | Best use | Key limitations |
| --- | --- | --- | --- | --- | --- | --- |
| Manual sideload, Office on the web | Manual | Current browser/user/document | XML manifest for Word/Excel/PowerPoint/OneNote, Outlook uses its own sideload path | Yes | Fast development in web clients | Browser-local; clearing cache or switching browsers requires sideload again; not tenant rollout |
| Manual sideload, desktop Office | Manual | Current desktop/user | XML manifest | Windows/Mac only for desktop Office | Desktop runtime debugging | Linux has no native Office desktop client; per-user/per-machine setup |
| Agents Toolkit unified package sideload | Automated CLI | Current developer account/client | Unified M365 package zip | Yes | Local developer install/update of the unified manifest package | It is still a sideload/debug install, not tenant assignment or broad rollout; requires `atk` auth and title ID cleanup |
| Centralized Deployment / Integrated Apps | Manual admin UI | Users, groups, or everyone | Office add-in / integrated-app package depending on portal path | Admin portal is web; PowerShell automation may not be Linux-friendly | Tenant rollout for Office add-ins | Propagation can take up to 24 hours; requires supported licenses, Exchange Online, Entra ID, and admin roles |
| Centralized Deployment PowerShell | Automated CLI | Users, groups, or everyone | Classic Office XML manifests | Not reliable in this Linux runtime | Scripted XML add-in deployment where cmdlets work | `ExchangeOnlineManagement` can authenticate but may not expose `New-OrganizationAddIn`; `O365CentralizedAddInDeployment` exposes cmdlets here but fails Linux auth with `kernel32.dll` |
| CLI for Microsoft 365 app catalog | Automated CLI | Tenant app catalog availability | Unified M365 package zip | Yes | Upload/update the unified package from Linux | Uploads to the app catalog lane; it is not the same as classic Office XML Centralized Deployment or guaranteed Office client assignment |
| Agents Toolkit / Teams Toolkit | Automated/dev CLI | Developer tenant/user, depending command | Unified M365 package zip | Yes | Developer install/provision/debug of unified package | More of a dev/provisioning lane than a replacement for tenant Centralized Deployment |
| Cloudflare dev tunnel + sideload | Automated local dev | Developer only | XML or unified package pointing at tunnel URL | Yes | Remote Office web testing from a local Vite server | Quick-tunnel host changes require Entra redirect and manifest regeneration; not for broad tenant rollout |

## Platform And Client Fit

| Client/platform | Classic XML add-in-only manifest | Unified M365 package manifest | Notes |
| --- | --- | --- | --- |
| Office on the web | Supported for Word, Excel, PowerPoint, OneNote; Outlook has a separate sideload path | Supported | Strongest cross-platform development target because it runs in the browser |
| Word/Excel/PowerPoint on Windows desktop | Supported on supported Microsoft 365/Office builds | Supported on Microsoft 365 subscription builds at or above the unified-manifest minimum | Classic XML remains the fallback for older Windows Office clients |
| Word/Excel/PowerPoint on Mac desktop | Supported on supported Office for Mac builds | Supported on recent Office for Mac builds at or above the unified-manifest minimum | Test both XML and unified package if supporting older Mac clients |
| Outlook classic on Windows | Supported | Supported on supported Microsoft 365 subscription builds | Outlook has separate manifest/runtime requirements from document surfaces |
| New Outlook on Windows | Limited to supported Outlook add-in model paths | Supported | Use Outlook-specific validation before relying on parity with classic Outlook |
| Outlook on Mac | Supported through Outlook add-in paths | Not supported for unified manifest according to Microsoft’s current unified-manifest table | Keep XML Outlook manifest for Mac coverage |
| Outlook mobile iOS/Android | Only Outlook mobile add-ins; Centralized Deployment supports Outlook mobile add-ins | Unified Office add-ins are not broadly supported on Office mobile | Do not assume Word/Excel/PPT mobile support |
| OneNote | Supported through the separate OneNote XML manifest | Unified manifest not supported | Keep `dist/package/development/onenote/onenote.manifest.xml` |
| Teams | Not an Office XML add-in host | Supported as part of a Teams/M365 app package | Teams app catalog upload is useful, but it does not replace Office XML add-in deployment |
| Linux desktop | No native Office desktop client target | No native Office desktop client target | Use Office on the web from Linux; CLI tooling can run on Linux |

## What To Run

### One-Shot Bootstrap

`bun bootstrap` is the umbrella flow:

```bash
bun bootstrap
```

Use the explicit lane aliases when you know the target:

| Command | Profile | Deployment lane | Intended use |
| --- | --- | --- | --- |
| `bun run bootstrap:dev` | `development` | `none` | Safe local bootstrap that packages and updates skills without tenant/catalog mutation |
| `bun run bootstrap:dev:sideload` | `development` | Agents Toolkit sideload | Build/package and install the unified dev package for the current developer |
| `bun run bootstrap:release:dry-run` | `internal-alpha-word-excel` | `catalog`, dry-run | Check release config, package path, and catalog command shape before mutating anything |
| `bun run bootstrap:release` | `internal-alpha-word-excel` | `catalog` | Production-like stable-origin release path for the current alpha profile |
| `bun run bootstrap:prod` | `internal-alpha-word-excel` | `catalog` | Alias for `bootstrap:release`; kept explicit so no one mistakes dev for prod |

It runs:

1. readiness checks,
2. repo-local Google WIF check/login for Gemini Enterprise skill provisioning,
3. manifest generation, validation, and packaging,
4. Gemini Enterprise skill upsert,
5. Entra SPA redirect sync for tenant/catalog deploys,
6. deployment/upload through the selected lane.

The release/catalog path is profile-aware. `bootstrap:release` packages
`dist/release/internal-alpha-word-excel-v<version>.zip`, syncs
`https://GE_ALPHA_WEB_DOMAIN/auth-redirect.html` into the release Entra app registration, then uploads
that release zip. It does not upload `development-m365-v<version>.zip`.

Deployment lane options:

| Bootstrap option | What it does | Works from Linux here? | Result |
| --- | --- | --- | --- |
| `--deployment-lane auto` | Prefer XML Centralized Deployment when the local PowerShell backend can run it; otherwise use unified app catalog upload | Yes | Chooses catalog upload in this runtime |
| `--deployment-lane catalog` | Upload/upsert `dist/release/development-m365-v<version>.zip` with CLI for Microsoft 365 | Yes | Unified package is uploaded/updated in the app catalog lane |
| `--deployment-lane xml` | Deploy/update classic XML manifests through Centralized Deployment PowerShell | No, not in this Linux runtime | Needs Windows PowerShell here unless Microsoft module support changes |
| `--deployment-lane none` or `--skip-tenant-deploy` | Package and update skills only | Yes | No tenant/app-catalog mutation |

For a Linux-native automated upload path:

```bash
bun bootstrap -- --deployment-lane catalog
```

For classic XML Centralized Deployment where the PowerShell backend works:

```bash
bun bootstrap -- --deployment-lane xml --tenant-backend Auto
```

Important: the catalog lane is an automated upload/update of the unified app package. It is not the
same as classic XML Centralized Deployment assignment to users/groups. If you need guaranteed Office
ribbon availability for classic XML add-ins across assigned users, use the XML Centralized
Deployment lane or the Integrated Apps portal.

Direct release checks:

```bash
bun run entra:sync:release -- --dry-run
M365_APP_PROFILE=internal-alpha-word-excel scripts/m365-app-catalog.sh upsert --dry-run
```

### Web Or Desktop Developer Sideload

Generate artifacts:

```bash
bun run setup:package
```

Use:

- `dist/package/development/xml/*.manifest.xml` for classic Office XML sideload.
- `dist/release/development-m365-v<version>.zip` for unified package upload/dev install.

Automated unified package sideload:

```bash
bun run sideload
```

The wrapper uses Agents Toolkit CLI:

```bash
atk install --file-path dist/release/development-m365-v<version>.zip
```

It records the returned title ID in `.ge-sideload/unified-sideload.json` when possible, so cleanup is:

```bash
bun run sideload:uninstall
```

If parsing fails, pass the title ID printed by `atk`:

```bash
bun run sideload:uninstall -- --title-id U_<title-id-guid>
```

### Tenant XML Deployment

Use this for classic Office add-in XML deployment:

```bash
bun run m365:addins
```

On this Linux runtime, live Centralized Deployment may be blocked by Microsoft module behavior. The
script can still package and dry-run:

```bash
M365_ADDIN_SKIP_PACKAGE=1 scripts/m365-tenant-addin.sh deploy --dry-run --upload-only --no-connect
```

If live deployment is blocked, run the printed PowerShell handoff on Windows or use the Integrated
Apps portal.

### Unified Package App Catalog Upload

Use this for the unified M365 package zip:

```bash
bun run m365:catalog
```

Direct commands:

```bash
M365_CLI_INSTALL=1 scripts/m365-app-catalog.sh login
M365_CLI_LOGIN=1 scripts/m365-app-catalog.sh upsert
```

Equivalent raw CLI shape:

```bash
m365 login --authType deviceCode
m365 teams app add --filePath dist/release/development-m365-v<version>.zip
```

Use `upsert` for repeat runs. It reads the `id` from the package `manifest.json`, tries
`m365 teams app get --id`, then runs `m365 teams app update` or `m365 teams app add`.

## Decision Table

| Goal | Recommended path |
| --- | --- |
| Test the task pane in Excel/Word/PPT/OneNote on the web | Manual web sideload with XML manifest |
| Test unified package behavior in M365/Teams app model | Unified package zip via Agents Toolkit or CLI for Microsoft 365 |
| Roll out classic Office add-in commands to users/groups | Centralized Deployment / Integrated Apps with XML manifests |
| Roll out from Linux without PowerShell module issues | Use CLI for Microsoft 365 for unified app catalog upload, or use admin portal; do not expect XML Centralized Deployment PowerShell to work in every Linux runtime |
| Support OneNote | Keep OneNote XML flow |
| Support older Office desktop clients | Keep XML manifests alongside unified package |
| Broad tenant deployment with stable URLs | Use stable HTTPS origin, regenerate manifests, then Integrated Apps/Centralized Deployment |
| Cloudflare quick tunnel debugging | Sideload only; never broad tenant rollout |

## Sources

- Microsoft: Office web sideloading notes distinguish classic add-in-only manifests from unified
  manifest sideloading and state browser-local storage behavior.
- Microsoft: unified manifest sideloading supports `atk install --file-path <zip>` and recommends
  uninstalling with the Agents Toolkit title ID returned by install.
- Microsoft: Centralized Deployment supports Windows, Mac, and web Office clients, and Outlook
  mobile add-ins; it requires supported licenses, Exchange Online, Entra ID, and admin permissions.
- Microsoft: Unified manifest enables one package for Microsoft 365 app distribution, but OneNote
  and Project do not support it; Microsoft’s platform table lists current web/Windows/Mac/Outlook
  support and unsupported mobile/perpetual/Outlook Mac cases.
- CLI for Microsoft 365: `m365 teams app add --filePath` uploads a Teams/M365 app package to the app
  catalog lane.
