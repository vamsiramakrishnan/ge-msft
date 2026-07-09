# Hosting Origin and Release Flow

This repo is client-direct. The hosted service is not the document worker and it is not the
Gemini Enterprise backend. It is the HTTPS origin that Office uses to download the add-in web app:
HTML, JavaScript, CSS, icons, and the MSAL redirect page.

## Runtime Model

```text
Microsoft 365 admin deployment or developer sideload
        |
        | installs a manifest/package
        v
+-------------------------------+
| Office host                   |
| Excel / Word / PowerPoint     |
| Outlook / OneNote / Teams     |
+-------------------------------+
        |
        | manifest points at:
        |   https://<host>/taskpane.html
        |   https://<host>/commands.html
        |   https://<host>/auth-redirect.html
        v
+----------------------------------------+
| Office embedded browser / WebView      |
|                                        |
| packages/web-shell React app           |
| Office.js bridge                       |
| MSAL / Entra auth                      |
| Gemini Enterprise StreamAssist client  |
+----------------------------------------+
        |                         |
        | Office.js               | HTTPS
        v                         v
Open Office document         Gemini Enterprise /
workbook / doc / deck /      Discovery Engine /
mail item / page / chat      StreamAssist APIs
```

The manifest and tenant deployment determine where the add-in appears. The hosted origin determines
where Office loads the browser app from. The browser app then does two separate jobs:

1. It calls Office.js inside the host to read, navigate, and mutate the current surface.
2. It calls Gemini Enterprise / Discovery Engine over HTTPS to stream plans, answers, command blocks,
   skills, sessions, and connector context.

That means a stable host is required even when all sensitive actuation remains client-side.

## Why A Stable HTTPS Origin Exists

Office will not run the add-in by reading files from this repository. It loads the add-in from URLs
declared in the manifest. The same origin also has to be registered in Entra for MSAL redirect flows.

| Concern | What points to the host |
| --- | --- |
| Task pane UI | `taskpane.html` in the Office manifest |
| Ribbon command page | `commands.html` in the Office manifest |
| Auth redirect | `auth-redirect.html` in Entra SPA redirect URIs and manifest runtime config |
| Static assets | JavaScript chunks, CSS, icons, and generated images |
| Cache/update behavior | Users reload the web app from this origin after the manifest has been installed |

Changing the host requires:

1. update `GE_DEV_WEB_ORIGIN` / release profile origin,
2. update `GE_DEV_WEB_DOMAIN`,
3. regenerate manifests,
4. update Entra redirect URIs,
5. repackage or redeploy the add-in package/manifests.

Cloudflare quick tunnels are useful for development because they expose a local Vite server over
HTTPS. They are not suitable for tenant rollout because the hostname is ephemeral.

## Hosting Choices

| Host | Best use | What we deploy | Advantages | Tradeoffs |
| --- | --- | --- | --- | --- |
| Cloudflare quick tunnel | Local development and sideloading | Nothing permanent; Vite runs locally | Fastest iteration from a remote workstation | Host changes frequently; every host change needs manifest and Entra redirect updates |
| Cloud Storage + Cloud CDN | Static production hosting | `packages/web-shell/dist-web` or `dist/package/<profile>/web` | Simple static asset hosting, good caching, no Node server | HTTPS with a custom domain normally needs load balancing/CDN setup; not a place for future server endpoints |
| App Engine Standard | Stable HTTPS host with minimal server code | A small Node static server plus built assets | Easy default HTTPS `appspot.com` origin, versions, traffic promotion, low operational overhead | App Engine app region is permanent; less flexible than Cloud Run for future APIs |
| Cloud Run | Stable HTTPS host plus future server endpoints | Container or source-deployed Node app plus built assets | Best path if we add tenant config, token brokerage, diagnostics, server-side policy, or proxy endpoints | Slightly more release plumbing than pure static hosting |

Current App Engine Standard guidance says the Node.js runtime is declared in `app.yaml` as
`runtime: nodejs VERSION`; as of July 8, 2026, the latest supported Node.js major is 24, so an App
Engine release should use `runtime: nodejs24` and a compatible `package.json` `engines.node` value.
App Engine automatically updates minor and patch releases within that major version, but not the
major version.

Cloud Run is the more general managed runtime. Google documents that it can deploy source for
Node.js and other supported frameworks, and the result runs in managed container instances. Use Cloud
Run if the hosted origin is likely to become more than static files.

Cloud Storage static hosting is valid for static assets, but HTTPS custom-domain serving usually
requires Google Cloud load balancing and certificate work. Use it when the add-in remains purely
static and the release pipeline already owns CDN setup.

## Recommended Release Shape

For development:

```text
Vite dev server
  -> Cloudflare quick tunnel
  -> write GE_DEV_WEB_ORIGIN / GE_DEV_WEB_DOMAIN
  -> update Entra redirect
  -> regenerate manifests/package
  -> sideload or dev-install
```

For production-like release:

```text
bun build web shell
  -> deploy web shell to stable HTTPS host
  -> write stable GE_DEV_WEB_ORIGIN / profile origin
  -> write stable GE_DEV_WEB_DOMAIN
  -> register https://<host>/auth-redirect.html in Entra
  -> regenerate manifests
  -> validate manifests
  -> package XML manifests and unified M365 zip
  -> deploy through chosen Microsoft lane
```

The Bun release wrapper makes that profile-aware:

```bash
bun run bootstrap:release:dry-run
bun run bootstrap:release
```

`bootstrap:release` currently uses the `internal-alpha-word-excel` profile. It reads
`GE_ALPHA_WEB_DOMAIN`, generates a manifest whose runtime URLs point at that host, syncs
`https://GE_ALPHA_WEB_DOMAIN/auth-redirect.html` into the Entra SPA redirect list, packages
`dist/release/internal-alpha-word-excel-v<version>.zip`, and upserts that package through the M365
catalog lane.

Microsoft deployment lane is separate from hosting lane:

| Microsoft lane | What it distributes | Still needs hosted origin? |
| --- | --- | --- |
| Office XML sideload | One XML manifest for a host | Yes |
| Centralized Deployment / Integrated Apps | XML add-in manifests or integrated app package | Yes |
| Unified M365 app catalog | `dist/release/development-m365-v<version>.zip` | Yes |
| Teams app catalog | Unified package app metadata | Yes |

## Choosing App Engine Or Cloud Run

Use App Engine Standard when the release goal is:

- stable HTTPS origin for Office manifests,
- no custom server behavior yet,
- simple versioned deployment and promotion,
- minimal operations.

Use Cloud Run when the release goal includes or will soon include:

- tenant config lookup,
- health and diagnostics endpoints,
- token or auth mediation,
- server-side allowlists or policy enforcement,
- request logging that cannot live in the browser,
- a shared API for all Office surfaces.

In both cases, the Office add-in still runs inside Office's embedded browser. The host only serves
and optionally assists that browser app.

## Release Guardrails

- Never broad-rollout a manifest that points to `*.trycloudflare.com`.
- Never publish a package with placeholder app IDs.
- Regenerate manifests after every host change.
- Register only the current stable `https://<host>/auth-redirect.html` redirect in Entra for the
  release app registration.
- Keep dev and production origins separate.
- Avoid `X-Frame-Options: DENY` or `SAMEORIGIN` on task pane pages; Office must be able to frame the
  add-in.
- Treat `dist/package/<profile>/web` as static deployable output, not source.

## Source References

- Google App Engine Standard Node.js runtime:
  <https://docs.cloud.google.com/appengine/docs/standard/nodejs/runtime>
- Google Cloud Run overview:
  <https://docs.cloud.google.com/run/docs/overview/what-is-cloud-run>
- Google Cloud Storage static website hosting:
  <https://docs.cloud.google.com/storage/docs/hosting-static-website>
