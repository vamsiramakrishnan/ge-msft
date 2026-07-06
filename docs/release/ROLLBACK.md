# Rollback Runbook

How to roll back a bad production (or alpha) release of Gemini Enterprise for Microsoft 365.
This file is also the artifact checked by the `rollback artifact availability` gate in
`tools/release/release-check.mjs` (`rollbackOutcome()`), which fails the release check when
`docs/release/ROLLBACK.md` is missing.

The architecture is client-direct (ADR-0001): there is no server tier to roll back. A release is
exactly three things, and each rolls back independently:

1. the **static web bundle** (`dist-web`) served from the production origin,
2. the **manifest package** uploaded to the Microsoft 365 admin center (plus the OneNote XML),
3. the **Entra / WIF configuration** the shell authenticates against.

Decide which of the three regressed before touching anything — most incidents need only #1.

---

## 1. Static-origin rollback (fastest, usually sufficient)

The manifest pins only the origin (`https://<GE_PROD_WEB_ORIGIN>`); everything the user runs is
fetched from that origin at task-pane load. Redeploying the previous `web/` tree from the previous
release's package directory rolls every user back on their next pane open — no admin-center
action, no user action.

```bash
# Identify the last-good release from the audit trail (see §4), then resync its web build.
gsutil -m rsync -r -d dist/package/production/web gs://<prod-bucket>
# Invalidate the CDN cache for the mutable entry points (see cache note below).
gcloud compute url-maps invalidate-cdn-cache <url-map-name> --path "/*" --async
```

**Cache behavior — read before invalidating selectively.** Vite emits content-hashed assets
(`assets/taskpane-<hash>.js`, `assets/taskpane-<hash>.css`), which are immutable: a previous
build's hashed assets never collide with the new build's, so they need no invalidation. The
exceptions are the **stable-path files**:

- `taskpane.html`, `commands.html`, `auth-redirect.html` — small HTML entry points,
- **`assets/commands.js`** — this is NOT content-hashed. The unified manifest references the
  stable path `assets/commands.js`, and the packagers copy the hashed `commands-<hash>.js` chunk
  onto that stable name (`tools/release/package-*.mjs` fail the build if the chunk is missing).
  Because the path is stable across releases, a CDN or browser can serve a **stale commands.js
  after rollback**. Serve it (and the HTML entry points) with `Cache-Control: no-cache` or a short
  max-age, and always invalidate those paths explicitly when rolling in either direction:

```bash
gcloud compute url-maps invalidate-cdn-cache <url-map-name> \
  --path "/assets/commands.js" --path "/taskpane.html" --path "/commands.html" \
  --path "/auth-redirect.html"
```

Also redeploy the previous `config/*.json` tenant-config files if the incident was a config push
(they are same-origin deploy artifacts, ADR-0009 — rolling back the bundle without the config, or
vice versa, is a valid partial rollback).

Verify: fetch `https://<origin>/taskpane.html` and `https://<origin>/assets/commands.js`, compare
their hashes against the previous release's `SHA256SUMS`/`walk` of the previous package dir, then
open the pane in Word on the web.

## 2. Manifest rollback (Microsoft 365 admin center)

Only needed when the *manifest itself* is bad (wrong id, wrong origin, broken ribbon/runtime
wiring) — a bad web bundle does not require this.

**Caveat — admins cannot downgrade a manifest version.** The admin center treats the app version
monotonically: re-uploading a package whose `version` is lower than (or equal to) the currently
deployed one is rejected or ignored by clients that already cached the newer version. Do not try
to re-upload last release's zip as-is. Use the **roll-forward-with-previous-bundle** pattern:

1. Check out the last-good tag; verify its package hash against that release's
   `production-artifact.json` (§4).
2. Bump `version` in the root `package.json` to a **new, higher** patch version (e.g. bad release
   was `1.4.0`, last good was `1.3.0` → release `1.4.1` with `1.3.0`'s content).
3. Regenerate and repackage with the production profile:
   `npm run manifests:generate:prod && npm run manifests:validate:prod && npm run package:prod`
   (the generated manifest version tracks `package.json`, so the new zip carries the higher
   version over the previous content).
4. Microsoft 365 admin center → **Settings → Integrated apps** → select the app → **Update** →
   upload the new zip. Assignments and consent are preserved on update.
5. Office clients pick up the manifest on their refresh cycle (web within hours, desktop up to
   24–72 h). The static origin remains the real kill switch in the meantime — see §1.

For the OneNote XML package, re-upload the regenerated `onenote.manifest.xml` through the same
Integrated apps flow (or the SharePoint app catalog if that is how it was deployed), with the same
higher-version rule (`<Version>` in the XML tracks `officeXmlVersion()`).

**Kill switch:** if the manifest is actively harmful, unassign users (Integrated apps → app →
**Users** → remove assignment) or block the app rather than waiting for a version roll-forward.

## 3. Entra / WIF configuration rollback

Identity config lives outside the artifact; it rolls back in the two consoles:

- **Entra app registration** (Microsoft Entra admin center → App registrations →
  `GE_PROD_ENTRA_CLIENT_ID`): revert changed SPA redirect URIs (must include
  `https://<origin>/auth-redirect.html`), delegated permission set (`openid`, `profile`, `email`,
  `User.Read`, plus any granted Graph read scopes), and re-run **admin consent** if permissions
  were reverted. The app must remain a public client / SPA — no client secret is ever added
  (client-direct, ADR-0001).
- **Google WIF** (Workforce Identity Federation pool/provider named by the shipped
  `VITE_WIF_POOL_ID` / `VITE_WIF_PROVIDER_ID`): revert provider attribute mappings/conditions and
  the allowed audience (the Entra client id). Never delete the pool to "roll back" — pools are
  soft-deleted and block re-creation under the same id for 30 days; edit the provider in place or
  disable it.
- If the incident was a rotated identifier baked into the bundle (`VITE_ENTRA_CLIENT_ID` etc.),
  that is a *bundle* rollback (§1) — the identifiers are public config compiled into the build.

After identity rollback, verify the full chain on one account: NAA sign-in in the pane → Entra
token → WIF STS exchange → a `streamAssist` round trip.

## 4. Audit trail: which bytes were live when

Every release records what shipped, so rollback targets are identified by hash, not by memory:

- `dist/release/production-artifact.json` (uploaded to the GitHub Release by
  `.github/workflows/release.yml`): profile, version, zip path + SHA-256, manifest SHA-256,
  OneNote manifest SHA-256, manifest version, web build path.
- `dist/release/SHA256SUMS`: hashes for the zip and generated manifests; verify with
  `npm run release:checksums` (runs `tools/release/checksums.mjs --verify`).
- `dist/release/release-status.json` / `release-status.md` (written by
  `npm run release:check`): the full gate table for the cut, including the live-host
  certification evidence that was matched (commit SHA + package hash + manifest version).

To confirm "what is currently deployed", hash the served files and match them against the
release's artifact json: `curl -s https://<origin>/taskpane.html | sha256sum`.

## 5. After any rollback

1. Re-run `npm run release:check` against the now-live tag; the report must PASS (or be BLOCKED
   only on external gates you can name).
2. Record the incident, the rolled-back version, and the hashes in the release notes of the
   roll-forward release.
3. New certification evidence (docs/release/HOST-CERTIFICATION.md) is required for the next cut —
   evidence is pinned to commit SHA and package hash, so it never carries over.
