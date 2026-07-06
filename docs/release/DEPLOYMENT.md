# Deployment Runbook

Production deployment of Gemini Enterprise for Microsoft 365. The release is client-direct
(ADR-0001): deploying means (a) putting the built web bundle on an HTTPS static origin,
(b) distributing the manifest package through the Microsoft 365 admin center, and (c) having a
per-tenant Entra app registration (plus the Google WIF pool/provider) that the bundle's baked
identifiers point at.

Inputs: the `production` profile artifacts from `npm run package:prod`
(`dist/package/production/{m365,onenote,web}` + `dist/release/production-m365-v<ver>.zip` +
`production-artifact.json` + `SHA256SUMS`), built in CI by `.github/workflows/release.yml`.

---

## 1. Static hosting for `dist-web`

The manifest pins `https://<GE_PROD_WEB_ORIGIN>`; every task pane, command runtime, and auth
redirect loads from that origin. Any static host with HTTPS + correct MIME types works. Two
supported options:

### Option A (default): GCS bucket + external HTTPS LB + Cloud CDN + custom domain

Keeps the origin in the same GCP project/region posture as the Gemini Enterprise resources
(residency pinning, CLAUDE.md constraint).

```bash
PROJECT=<gcp-project> ; REGION=<pinned-region> ; DOMAIN=addin.contoso-ge.com
BUCKET=ge-m365-prod-web

# 1. Bucket (uniform access, public read of objects via IAM, not legacy ACLs)
gcloud storage buckets create gs://$BUCKET --project=$PROJECT --location=$REGION \
  --uniform-bucket-level-access
gcloud storage buckets add-iam-policy-binding gs://$BUCKET \
  --member=allUsers --role=roles/storage.objectViewer

# 2. Upload the packaged web tree (from the exact release package, not a rebuild)
gsutil -m rsync -r -d dist/package/production/web gs://$BUCKET

# 3. Cache headers: hashed assets immutable; stable entry points + commands.js revalidate.
#    assets/commands.js is NOT content-hashed (the packagers copy the hashed chunk onto the
#    stable path the manifest references) — treat it like HTML, or rollback goes stale
#    (see docs/release/ROLLBACK.md §1).
gsutil -m setmeta -h "Cache-Control:public, max-age=31536000, immutable" "gs://$BUCKET/assets/*"
gsutil setmeta -h "Cache-Control:no-cache" \
  gs://$BUCKET/taskpane.html gs://$BUCKET/commands.html gs://$BUCKET/auth-redirect.html \
  gs://$BUCKET/assets/commands.js

# 4. Load balancer + Cloud CDN in front of the bucket
gcloud compute backend-buckets create ge-m365-web-backend \
  --gcs-bucket-name=$BUCKET --enable-cdn --cache-mode=USE_ORIGIN_HEADERS
gcloud compute url-maps create ge-m365-web-map --default-backend-bucket=ge-m365-web-backend

# 5. Google-managed HTTPS certificate for the custom domain
gcloud compute ssl-certificates create ge-m365-web-cert --domains=$DOMAIN --global
gcloud compute target-https-proxies create ge-m365-web-proxy \
  --url-map=ge-m365-web-map --ssl-certificates=ge-m365-web-cert
gcloud compute addresses create ge-m365-web-ip --global
gcloud compute forwarding-rules create ge-m365-web-https \
  --address=ge-m365-web-ip --global --target-https-proxy=ge-m365-web-proxy --ports=443

# 6. DNS: point $DOMAIN (A/AAAA) at the reserved address; the managed cert activates
#    after DNS resolves (can take ~15-60 min). Verify:
gcloud compute ssl-certificates describe ge-m365-web-cert --global --format='value(managed.status)'
curl -sI https://$DOMAIN/taskpane.html | head -5
```

Office add-in hosting notes: everything is same-origin so no CORS config is needed for the shell
itself; do not add `X-Frame-Options: DENY` or a `frame-ancestors` that excludes Office hosts — the
pane runs in an embedded browser/iframe on the web hosts.

### Option B (alternative): Cloud Run + nginx

If an LB + bucket is heavier than wanted, serve the same `web/` tree from a minimal nginx
container on Cloud Run (`--region` pinned per tenant residency), with the same Cache-Control
split (immutable `assets/*`, `no-cache` for the HTML entry points and `assets/commands.js`) baked
into `nginx.conf`, and a domain mapping or LB for the custom domain. This is a transparent static
server only — it holds no credentials and adds no API surface (it is not the retired gateway).

### The `config/*.json` tenant-config directory (ADR-0009)

The shipped product envelope includes a `config/` directory of JSON documents served from the
**same origin** as the bundle, alongside it (e.g. `https://$DOMAIN/config/<name>.json`):

- Same-origin is a security property, not a convenience: ADR-0009 forbids tenant-editable
  absolute URLs (an arbitrary `proxyUrl` receiving a federated bearer token is an exfiltration
  primitive). The shell resolves config/proxy references through a build-time allowlist or
  same-origin policy only.
- These files can **select, route, reduce, or disable within the shipped envelope** — they can
  never mint permissions, change the release profile, or point at new origins. The release
  profile is a compiled constant of the artifact.
- Deploy them with the bundle (`gsutil rsync` above carries them if present under `web/config/`),
  give them `Cache-Control: no-cache` so config rolls take effect on next pane load, and treat a
  config-only change as a deploy (record hashes; rollback per ROLLBACK.md §1).
- Tenant-editable *ring* configuration additionally lives in SharePoint per ADR-0009
  (`/Shared Documents/config/config.json`, read via Graph with Selected permissions); the
  same-origin `config/` directory is the build-time/product side of that split.

## 2. Distribution without sideloading (Microsoft 365 admin center)

Sideloading is a dev workflow. Tenant-wide distribution is centralized deployment via Integrated
apps — users get the add-in with no local action:

1. Sign in to the **Microsoft 365 admin center** (`admin.microsoft.com`) as Global Admin or an
   admin role with app management rights.
2. **Settings → Integrated apps → Upload custom apps.**
3. App type: **Office Add-in / app built with the Teams Toolkit (unified manifest)**; upload
   `dist/release/production-m365-v<ver>.zip` (the zip contains `manifest.json` + icons — this is
   the admin-center artifact; the `web/` tree is *not* uploaded, it is served from the origin).
4. **Assign users**: start with a pilot group (ring 0/1 per ADR-0009), not "Entire organization";
   widen by group as certification and telemetry allow.
5. **Accept permissions / admin consent**: the wizard surfaces the `webApplicationInfo` Entra app
   (GE_PROD_ENTRA_CLIENT_ID) and, for Teams, the resource-specific consent
   (`OnlineMeetingTranscript.Read.Chat`). Accepting grants tenant-wide admin consent for the
   delegated scopes (§3). If skipped here, grant consent from the Entra portal instead.
6. **Deploy**. Propagation to Office clients is not instant: web hosts typically within hours,
   desktop ribbons up to 24–72 h.
7. Verify with a pilot user: Word on the web → Home ribbon → Gemini group appears without any
   sideload; network panel shows assets from the production origin.

**Teams admin center note:** the same unified package carries the Teams tab/bot/message-extension
surfaces. If the tenant manages Teams apps separately, also check **Teams admin center →
Teams apps → Manage apps**: the uploaded app must be Allowed, org-wide app settings must permit
custom apps, and the app can be pinned via setup policies. Blocking the app there hides only the
Teams surfaces, not the Office ones.

**OneNote (Package B):** OneNote still uses the legacy XML manifest
(`dist/package/production/onenote/onenote.manifest.xml`). Deploy it through the same admin center
Integrated apps flow (**Upload custom apps → Office Add-in → upload manifest file**) or, where the
tenant uses one, the SharePoint app catalog for Office add-ins. It is a separate app entry with
its own app id (GE_PROD_ONENOTE_APP_ID) and its own assignment list.

## 3. Per-tenant Entra app registration

Each customer tenant registers its own public-client SPA app; its client id becomes that tenant's
`GE_PROD_ENTRA_CLIENT_ID` / baked `VITE_ENTRA_CLIENT_ID`. Manual steps (Entra admin center →
App registrations):

1. **New registration** — name e.g. `Gemini Enterprise for Microsoft 365`; supported account
   types: **single tenant** (this org only).
2. **Authentication → Add a platform → Single-page application** with redirect URI
   `https://<origin>/auth-redirect.html` (the packaged auth redirect page). Enable no implicit
   grant; PKCE/auth-code is the SPA default. Add **no** client secret or certificate — the app
   must remain a public client (client-direct; a secret would be a credential we forbid clients
   to hold).
3. Also enable **Allow public client flows: No** (SPA + NAA needs no ROPC/device-code) and, for
   NAA in Office, make sure the registration includes the broker-compatible SPA redirect
   (`brokeredRedirectUri` support comes from the SPA platform entry above).
4. **API permissions** — delegated Microsoft Graph: `openid`, `profile`, `email`, `User.Read`.
   Add further delegated read scopes only as features require (prefer `Sites.Selected` over
   all-sites, per the repo's least-privilege constraint). No application permissions.
5. **Grant admin consent** for the tenant (button on the API permissions page), or rely on the
   Integrated-apps deployment consent step (§2.5).
6. **Expose an API** (for the unified manifest's `webApplicationInfo.resource`): set the
   Application ID URI to `api://<webDomain>/<client-id>` and add the `access_as_user` scope
   consented for the Office/Teams client ids — required for Teams SSO/NAA.
7. Record the client id + tenant id; they feed the manifest generation (`GE_PROD_ENTRA_CLIENT_ID`)
   and the web build (`VITE_ENTRA_CLIENT_ID`, `VITE_ENTRA_TENANT_ID`), and the Google WIF provider
   audience for that tenant.

### Script sketch (illustrative — review before running; ids/scopes must match your tenant)

```bash
# --- az CLI sketch: single-tenant SPA public client for the add-in -----------------
# ILLUSTRATIVE ONLY: verify each step against your tenant's policies before use.
ORIGIN="https://addin.contoso-ge.com"
APP_NAME="Gemini Enterprise for Microsoft 365"

APP_ID=$(az ad app create --display-name "$APP_NAME" \
  --sign-in-audience AzureADMyOrg \
  --query appId -o tsv)

# SPA platform + redirect (az has no --spa-redirect-uris flag; patch via Graph)
az rest --method PATCH \
  --url "https://graph.microsoft.com/v1.0/applications(appId='$APP_ID')" \
  --body "{\"spa\":{\"redirectUris\":[\"$ORIGIN/auth-redirect.html\"]}}"

# Delegated Graph permissions: openid, profile, email, User.Read
GRAPH=00000003-0000-0000-c000-000000000000
az ad app permission add --id "$APP_ID" --api $GRAPH --api-permissions \
  37f7f235-527c-4136-accd-4a02d197296e=Scope \
  14dad69e-099b-42c9-810b-d002981feec1=Scope \
  64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0=Scope \
  e1fe6dd8-ba31-4d61-89e7-88639da4683d=Scope
# (openid, profile, email, User.Read respectively)

# Service principal + tenant-wide admin consent
az ad sp create --id "$APP_ID"
az ad app permission admin-consent --id "$APP_ID"

echo "GE_PROD_ENTRA_CLIENT_ID=$APP_ID"
```

The same registration can be created via Microsoft Graph directly
(`POST https://graph.microsoft.com/v1.0/applications` with `signInAudience: "AzureADMyOrg"` and a
`spa.redirectUris` block, then `POST /servicePrincipals` and an `oauth2PermissionGrants` entry for
the consent) — same shape, same caveat: **the sketch is illustrative, not a hardened provisioning
tool**.

After registration, wire the Google side: add/confirm the tenant's WIF workforce pool provider
trusts this client id as audience, then run the §2 deployment and a live-host certification pass
(docs/release/HOST-CERTIFICATION.md) before widening assignment.

## 4. Order of operations for a cut

1. Tag `v<version>` (must equal root `package.json` version) → `.github/workflows/release.yml`
   gates (typecheck/lint/test/build), generates + validates production manifests, packages, and
   publishes the zip + `production-artifact.json` + `SHA256SUMS` to the GitHub Release.
2. Deploy `web/` (+ `config/`) to the static origin (§1); verify served hashes against
   `SHA256SUMS`.
3. Live-host certification against the exact package (docs/release/HOST-CERTIFICATION.md);
   evidence lands in `evidence/host-certification/`.
4. Upload/update the package in the admin center (§2) to the pilot ring; widen per plan.
5. Keep `docs/release/ROLLBACK.md` at hand; the audit trail for "what shipped" is the release's
   `production-artifact.json`.
