# Development Server and Cloudflare Tunnel

## Why a Server Is Still Required

The add-in is client-direct, but Office still loads the task pane, commands page, icons, and redirect
page from an HTTPS web origin. "Client-direct" means:

- The browser calls Entra, Google STS, and Gemini Enterprise directly.
- No gateway stores client secrets or durable tokens.
- The dev server only serves static web assets and Vite development bundles.

It does not mean the add-in can run from local files. Office web and desktop clients require a web
origin listed in the manifest.

## Local Dev Server

Use port `13000` for the remote workstation flow:

```bash
GE_DEV_PORT=13000 npm run dev -w packages/web-shell -- --host 0.0.0.0 --port 13000
```

For local desktop-only testing you can use localhost HTTPS, but a remote browser or Office on the web
needs a public HTTPS origin.

## Cloudflare Quick Tunnel

Start a tunnel that forwards to the Vite server:

```bash
cloudflared tunnel --url https://localhost:13000
```

Use the generated `https://<name>.trycloudflare.com` origin as:

```bash
GE_DEV_WEB_ORIGIN=https://<name>.trycloudflare.com
GE_DEV_WEB_DOMAIN=<name>.trycloudflare.com
```

Then regenerate manifests.

Important:

- Keep the Vite server running while the tunnel is running.
- If the tunnel returns 502, the local dev server is down, on the wrong port, or not reachable by
  `cloudflared`.
- Add `https://<name>.trycloudflare.com/auth-redirect.html` as an Entra SPA redirect URI.
- Re-upload or refresh the sideloaded manifest whenever the origin changes.

## Cloud Workstations Hostnames

The release tooling can derive a dev origin when:

```bash
GOOGLE_CLOUD_WORKSTATIONS=true
WEB_HOST=<workstation-host>
GE_DEV_PORT=13000
```

It becomes:

```text
https://13000-<workstation-host>
```

Use `GE_DEV_WEB_ORIGIN` explicitly when a Cloudflare tunnel is clearer.
