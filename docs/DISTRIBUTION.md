# Distribution: CDN hosting, sideloading, and the Microsoft marketplace

How the web shell (the add-in's entire client) gets from this repo onto a user's screen. Three
stages, each strictly more public than the last:

1. **Dev sideload** — cloudflared tunnel + ATK install. Already built: `bun run sideload`.
2. **CDN deploy** — a static bundle on real HTTPS hosting, pointed at by your manifests.
3. **Marketplace publish** — AppSource / Microsoft 365 admin-center distribution.

---

## 1. What an add-in actually deploys

The add-in is **static files + a manifest**:

| Piece | Where it lives | What points at it |
| --- | --- | --- |
| Hosting status page and Teams launch forwarder | `packages/web-shell/dist-web/index.html` | Hosting root `/` and Teams `/?host=teams` |
| Task pane, commands runtime, functions runtime, auth redirect | `packages/web-shell/dist-web/` (Vite build) | manifest `SourceLocation` / `FunctionFile` URLs |
| `functions.json` (custom-function metadata) | `dist-web/functions.json` | Excel XML manifest `<ExtendedOverrides>` |
| Manifests | `dist/package/<profile>/` via `bun run package:dev` / `package:alpha` | uploaded to M365 |

There is **no server** to deploy (client-direct, ADR-0001): the browser federates identity and
calls Gemini Enterprise directly. Hosting is therefore any static HTTPS origin with:
- a **valid publicly-trusted TLS certificate** (Office refuses self-signed outside dev),
- permissive-enough CORS for the hosts that fetch assets (the Office host iframes the page),
- stable URLs — they are baked into manifests users already installed.

**Config is build-time.** All `VITE_GE_*` values are inlined by Vite from `.env` at build time.
A deployed bundle is tenant-pinned until rebuilt. For one tenant per origin this is correct and
leak-proof; do not try to make one CDN bundle serve many tenants.

## 2. Build the deployable bundle

```bash
# production env first (origins, assistant path, WIF config), then:
bun run release:web
```

The script fails closed on the things that ruin deployments: missing entry pages,
`REPLACE_*` template tokens left in the bundle, `localhost` origins baked into JS, and any
secret findings (`tools/release/secret-scan.mjs` covers `dist-web`).

## 3. Host it (pick one)

### Cloudflare Pages (matches the existing cloudflared dev setup)

```bash
cd packages/web-shell && bunx wrangler pages deploy dist-web --project-name ge-msft-web
```
Then add the Pages domain to Entra → App registration → Authentication → SPA redirect URIs
(`https://<project>.pages.dev/auth-redirect.html`) via `bun run entra:sync:release -- --origin …`,
and regenerate manifests with the same origin.

### Firebase Hosting

```bash
firebase init hosting   # public dir: packages/web-shell/dist-web; SPA rewrite not needed
bun run deploy:web      # rebuilds, checks, then deploys Hosting
```

The generated `index.html` makes the hosting root a valid readiness page and forwards
`/?host=teams` to the task-pane shell. Do not replace it with a catch-all SPA rewrite.

### Azure Static Web Apps

```bash
az staticwebapp create -n ge-msft-web -g <rg> \
  --source https://github.com/<org>/<repo> --app-location packages/web-shell/dist-web
```

### Any object store + CDN (S3/GCS + CloudFront/Cloud CDN)

Upload `dist-web/**` with a long-cache policy for `/assets/*` and `no-cache` for the HTML
entry points and `functions.json` (they must re-resolve on each deploy).

After hosting: **update both manifest families** to the new origin
(`tools/release/common.mjs` profiles → `bun run manifests:generate && bun run validate:manifests`),
then redistribute the manifests (stage 4 below).

## 4. Sideloading (fast paths)

The guided flow is `bun run sideload` (ATK login, tunnel, package, install, state tracking in
`.ge-sideload/`). When you only have a zip and a browser:

| Host | Steps |
| --- | --- |
| **Microsoft 365 web** | office.com → any app (e.g. Word) → sidebar "Add-ins" / All apps → **Upload My Add-in** → choose `dist/package/development/m365-unified.zip` (or alpha) |
| **Outlook web** | outlook.office.com → ⚙ Settings → **Integrate all apps / Custom apps** → upload the Outlook XML manifest (`dist/manifests/outlook.xml`) |
| **Windows desktop** | Insert → My Add-ins / Get Add-ins → **More Add-ins** → My Add-ins → Upload, or run `scripts/m365-tenant-addin.ps1` for a shared-folder catalog |
| **OneNote** | OneNote web only, separate legacy package: upload `manifests/onenote.manifest.xml` variant from `dist/manifests/` |
| **Teams** | Teams → Apps → **Manage your apps → Upload an app** (custom app upload must be enabled by admin policy) |

Admin-center route (no end-user action): Microsoft 365 admin center → Settings → Integrated apps →
**Upload custom apps** → **Office Add-in**. Upload
`dist/package/development/centralized/office.manifest.xml` for Word/Excel/PowerPoint, then upload
`dist/package/development/centralized/outlook.manifest.xml` separately for Outlook, and assign the
same pilot user/group to both. This classic XML lane is the default for Mac desktop pilots. Keep the
unified zip in the Agents Toolkit/M365 catalog development lane.

## 5. Marketplace (AppSource / Microsoft 365 Store) checklist

Publishing goes through **Partner Center** (Partner Center → Windows & Xbox → Microsoft 365
program) which forwards to AppSource review. Nothing here can be automated away — these are
account/review requirements:

- [ ] Partner Center **Microsoft 365 program** enrollment (company verification, tax/payout).
- [ ] Replace every `REPLACE_*` placeholder in `manifests/m365-unified.manifest.json`:
      app GUID, Entra client id, web domain, developer name/urls.
- [ ] **Owned, validated domains**: privacy policy, support, EULA/terms URLs must be live on a
      domain you control; `validDomains` must match them exactly (wildcards rejected).
- [ ] Icons at 64×64 + color/outline variants per spec; no placeholder art anywhere in UX
      (review rejects lorem/TODO states).
- [ ] **SSO/Graph consent story documented for reviewers** with a test tenant + test accounts
      (client-direct uses Entra NAA — include the WIF federation explanation; reviewers reject
      unexplainable token flows).
- [ ] Version discipline: semver the manifest + `versionOverrides`; keep `dist/release`
      checksums (`bun run release:checksums`) as submission evidence.
- [ ] Both packages submitted separately if shipping OneNote (Package B uses the legacy XML
      schema; note this in submission notes).
- [ ] Validation policies worth pre-checking: no auto-send of mail (we never send — good),
      explicit user consent for writes (our preview/approve gate), working "get started"
      experience on first open.

Realistic sequencing: **admin-center custom-app deployment for pilots now; AppSource after two
design-partner tenants**, since review turnaround (days–weeks) punishes iteration.
