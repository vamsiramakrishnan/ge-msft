# How to run and deploy the add-in

All commands run on your own machine, from the repo root:
```bash
cd ge-msft
```

---

## 1. Quick developer workflow

Follow these steps to set up, develop locally, build, and package the add-in:

### 1. Install packages
```bash
bun install
```

### 2. Set up config (first time only)
```bash
cp packages/web-shell/.env.example packages/web-shell/.env
# then fill in the 7 required VITE_* values, plus GE_DEV_WEB_ORIGIN and the GE_DEV_*_APP_ID GUIDs
```
*(See [Section 5: Config reference](#5-config-packagesweb-shellenv) below for exact details on required variables).*

### 3. Check the UI
```bash
bun run dev                              # UI preview with sample data, no sign-in needed
bun run --filter @ge/web-shell dev       # the real task pane, at https://localhost:13000
```
Stop either one with `Ctrl+C`.

### 4. Build the web app
```bash
bun run --filter @ge/web-shell build     # quick build → packages/web-shell/dist-web/
bun run release:web                      # release build: same output, plus safety checks
```
> **Note:** Use `release:web` for anything you deploy. It fails if the bundle still contains placeholders, localhost URLs, or secrets.

### 5. Create the manifests
```bash
bun run manifests:generate --profile development   # → dist/manifests/
bun run manifests:validate --profile development   # check them
bun run package:dev                                # zips → dist/release/*.zip
```
For the production manifest, use `--profile production` and `bun run package:prod` instead (see
[section 6](#6-production-manifest-the-production-profile)).

### 6. Optional: full checks before sharing
```bash
bun run typecheck && bun run test && bun run lint
```

---

## 2. Deployment

`bun run release:web` writes the production web app to `packages/web-shell/dist-web/`. It fails if a
localhost URL, a placeholder or a secret is left in the build. Host that folder on any HTTPS static
host (see `setup/08-hosting-origin-and-release.md`). Serve HTML with no caching, cache the hashed
JS/CSS long-term, and do not set `X-Frame-Options`: Office shows the task pane in an iframe.

---

## 3. The architecture in one minute

An Office add-in is a **web page that Office opens in a side panel**. To make it work, four components must agree on **one web address** (`<your-web-app>`, without `https://`):

| Piece | What it is | Where it lives |
| --- | --- | --- |
| **Web app** | The panel's HTML/JS files | Private Cloud Run → `https://<your-web-app>` |
| **Manifest (XML)** | Tells Office where to place the button and what URL to open | `dist/manifests/` or `dist/release/*.zip` (uploaded to Office) |
| **Entra app registration** | Handles user sign-in via Microsoft account; directs back to registered URLs | Azure portal → App registrations (`VITE_ENTRA_CLIENT_ID`) |
| **Config (`.env`)** | Discovery Engine, identity pools, and origins baked into the build | `packages/web-shell/.env` |

**The One Rule:**
```
Web app origin == Address in the manifest == Address registered in Entra
```
If any one of these differs, the task pane will fail to load or sign-in will fail.

There is **no custom backend server** to maintain. After loading the static assets, the task pane communicates directly with Microsoft (MSAL authentication) and Google (Gemini Enterprise APIs via Workforce Identity Federation).

---

## 4. Register the address in Entra (once per address)

The Entra app registration must list **all three** links under **Single-page application (SPA)** redirect URIs:

```
https://<your-web-app>/auth-redirect.html
brk-multihub://<your-web-app>
brk-9199bf20-a13f-4107-85dc-02114787ef48://<your-web-app>
```

### Why three redirect URIs:
- `https://<your-web-app>/auth-redirect.html` is the standard browser/popup fallback redirect.
- The two `brk-…` links enable **Nested App Authentication (NAA)**. Office authenticates the user directly through its built-in broker. **Outlook requires NAA with these exact schemes**—without them, Outlook sign-in will fail.
- Note: `9199bf20-…` is Microsoft's fixed broker client ID, **not** your app's client ID. Keep that GUID as-is; only replace `<your-web-app>` with your domain.

### How to configure in Azure Portal:
1. Go to Azure Portal → **App registrations** → Select the add-in app (`VITE_ENTRA_CLIENT_ID`).
2. Select **Manifest** (the JSON editor), as the standard Authentication UI may reject custom `brk-` URI schemes.
3. Under `"spa"` → `"redirectUris"`, ensure all three links are present, then save.

---

## 5. Config: `packages/web-shell/.env`

Create `packages/web-shell/.env` from `.env.example`:

```bash
cp packages/web-shell/.env.example packages/web-shell/.env
```

### Required Variables:
1. **7 Core `VITE_*` Configuration Values**:
   - `VITE_GCP_PROJECT`: GCP Project ID hosting Gemini Enterprise / Discovery Engine.
   - `VITE_GCP_LOCATION`: Location matched to data-residency (e.g. `eu`, `us`, or region).
   - `VITE_GE_ENGINE`: Gemini Enterprise engine/app ID.
   - `VITE_WIF_POOL_ID`: Workforce Identity Federation pool ID.
   - `VITE_WIF_PROVIDER_ID`: Workforce Identity Federation provider ID.
   - `VITE_ENTRA_CLIENT_ID`: Entra Application (client) GUID.
   - `VITE_WIF_ID_TOKEN_SCOPES`: Scopes for token exchange (`openid profile email User.Read`).
2. **Web Origin & Domain**:
   - `GE_DEV_WEB_ORIGIN`: Base HTTPS URL where the web app is hosted (e.g. `https://<your-web-app>`).
   - `GE_DEV_WEB_DOMAIN`: Hostname without protocol (e.g. `<your-web-app>`).
3. **Application GUIDs (`GE_DEV_*_APP_ID`)**:
   - `GE_DEV_APP_ID`: Unified Microsoft 365 development package GUID.
   - `GE_DEV_OFFICE_XML_APP_ID`: Centralized Word + Excel + PowerPoint XML add-in GUID.
   - `GE_DEV_OUTLOOK_APP_ID`: Centralized Outlook XML add-in GUID.
   - `GE_DEV_ONENOTE_APP_ID`: Separate OneNote XML add-in GUID.

---

## 6. Production manifest: the `production` profile

The **development** profile is for testing: it's named "ContosoAI Dev", includes every app, and accepts
placeholder IDs and localhost. The **`production`** profile builds the manifest real users get:

| | Development (`--profile development`) | Production (`--profile production`) |
| --- | --- | --- |
| Name in Office | ContosoAI Dev | ContosoAI |
| Apps | Word, Excel, PowerPoint, Outlook, OneNote | Only those in `GE_PROD_SURFACES` (default Word, Excel, PowerPoint, Outlook) |
| App IDs | `GE_DEV_*_APP_ID` (placeholders allowed) | `GE_PROD_*_APP_ID` (real, permanent, required) |
| Checks | Loose | Fails on placeholder or reused IDs, localhost, IP addresses, `http://`, missing values |
| Output | `dist/package/development/` | `dist/package/production/`, `dist/release/production-*.zip` |

The two use **different app IDs**, so "ContosoAI Dev" and "ContosoAI" can be installed side by side.

### Settings to add

Put these in `packages/web-shell/.env` (on a laptop) or in the CI environment (in CI). Exported
variables override `.env`.

| Setting | Required | Example | Notes |
| --- | --- | --- | --- |
| `GE_PROD_WEB_DOMAIN` | Yes | `addin.example.com` | Host of the production web app, no `https://`. Must be the address in Entra (section 4) |
| `GE_PROD_APP_ID` | Yes | `a1b2…` (a GUID) | Identity of the unified package (`m365/`) |
| `GE_PROD_OFFICE_XML_APP_ID` | Yes | GUID | Identity of `centralized/office.manifest.xml` (Word, Excel, PowerPoint) |
| `GE_PROD_OUTLOOK_APP_ID` | Yes | GUID | Identity of `centralized/outlook.manifest.xml` |
| `GE_PROD_DEVELOPER_NAME` | Yes | `PT Bank Contoso Tbk` | Publisher shown in Office and the admin center |
| `GE_PROD_WEBSITE_URL` | Yes | `https://…` | Real organisation page, HTTPS only |
| `GE_PROD_PRIVACY_URL` | Yes | `https://…` | Real privacy page, HTTPS only |
| `GE_PROD_TERMS_URL` | Yes | `https://…` | Real terms page, HTTPS only |
| `GE_PROD_SUPPORT_URL` | Yes | `https://…` | Real help/support page, HTTPS only |
| `GE_PROD_ENTRA_CLIENT_ID` | No | GUID | Defaults to `VITE_ENTRA_CLIENT_ID`. Leave it unset unless production uses a different Entra app |
| `GE_PROD_SURFACES` | No | `word,excel,powerpoint,outlook` | Which apps to include. Add `onenote` later (see below) |
| `GE_PROD_ONENOTE_APP_ID` | Only with OneNote | GUID | Identity of the separate OneNote manifest |

### App IDs: how to handle them

An app ID is how Office recognizes the add-in. It's **not** a secret, and it's **not** the Entra client
ID. It's an identity you create yourself.

1. **Generate each one once:**
   ```bash
   uuidgen | tr A-Z a-z
   ```
   Run it once per ID: three IDs, or four with OneNote.
2. **Never change them.** If an ID changes, Office treats the add-in as brand new: users get a second
   copy, and the old one has to be removed by hand.
3. **Keep them all different,** from each other and from the `GE_DEV_*` IDs. The build fails if two
   match.
4. **Store them where they won't be lost.** `.env` isn't committed to Git, so record the production
   IDs in the team's secrets store or the CI variables as well. A new laptop or CI runner must use the
   **same** IDs.
5. **Changing the domain later doesn't change the IDs.** Keep the IDs; only regenerate the manifest.

### Commands

```bash
bun run release:web                                          # build the web app first
bun run manifests:generate --profile production         # → dist/manifests/production.*
bun run manifests:validate --profile production         # checks every production rule
bun run package:prod                                         # → dist/release/production-*.zip
```

Each run first deletes that profile's previous files, so a failed run never leaves old manifests
behind. If a value is missing or invalid, the command stops and names it (for example
`Missing production manifest configuration: GE_PROD_APP_ID`).

### What to upload

| File | Upload where | Covers |
| --- | --- | --- |
| `dist/release/production-m365-v<version>.zip` | M365 admin center → Integrated apps (one upload) | All the apps in `GE_PROD_SURFACES` |
| `dist/package/production/centralized/office.manifest.xml` and `outlook.manifest.xml` | M365 admin center → Office add-in → upload `.xml` | Word, Excel and PowerPoint in one file; Outlook in its own |

Use **one** of the two, not both. Before the first production upload:
- the production domain has all three Entra redirect URIs (section 4);
- the web app at that domain has been built with the production `VITE_*` values.

### Adding OneNote later

Set `GE_PROD_SURFACES=word,excel,powerpoint,outlook,onenote` and add a new `GE_PROD_ONENOTE_APP_ID`.
The build then also writes `dist/package/production/onenote/onenote.manifest.xml`. OneNote only
accepts the classic XML format, so it's uploaded separately.

---

## 7. Install and test the add-in

Before uploading, remove any previously uploaded copies to avoid GUID collisions in Office cache.

### Option 1: Sideload for individual testing
After running `bun run manifests:generate --profile development`:

- **Word / Excel / PowerPoint (Web)**:
  1. Open a document on [office.com](https://office.com).
  2. Navigate to **Home** → **Add-ins** → **More Add-ins** → **My Add-ins** → **Upload My Add-in**.
  3. Select `dist/manifests/word.xml`, `excel.xml`, or `powerpoint.xml`.
  4. Launch **Gemini Enterprise** from the ribbon.

- **Outlook (Web and new Outlook)**:
  1. Ensure the `brk-` redirect URIs are configured in Entra first (section 4).
  2. Open [aka.ms/olksideload](https://aka.ms/olksideload). It opens the **Add-ins for Outlook** dialog directly.
  3. Select **My add-ins**, scroll to the **bottom** to **Custom Addins**, then **Add a custom add-in** → **Add from file...**
  4. Select `dist/package/development/centralized/outlook.manifest.xml` (the same file as
     `dist/manifests/development.outlook.manifest.xml`).
  5. Open an email (or start a new one) and launch the add-in from the message toolbar.

  Deploying the web app, or uploading the manifest in the admin center, does **not** add the add-in to
  your own Outlook straight away. Until it's added under **My add-ins** (or an admin deployment has
  propagated), it doesn't appear in the **Add-ins for Outlook** list.

### Option 2: Centralized deployment (M365 Admin Center)
For rolling out to groups or the entire tenant:
1. Navigate to the [Integrated Apps Center](https://admin.cloud.microsoft/?#/Settings/IntegratedApps).
2. Choose **Upload custom apps** → **Office Add-in** → **Upload manifest file (.xml) from device**.
3. Upload `dist/manifests/office.xml` (covers Word, Excel, PowerPoint) and `dist/manifests/outlook.xml` (Outlook).
4. Assign target users or security groups (deployment can take up to 24 hours to propagate).

---

## 8. What to redo when things change

| What changed | Actions required |
| --- | --- |
| Code only | `bun run release:web` → CI/CD deploys to Cloud Run |
| `.env` values (`VITE_*`) | `bun run release:web` → CI/CD deploys to Cloud Run (values are baked in at build time) |
| Web origin / domain | Rebuild web app → Regenerate manifests (`manifests:generate`) → Update Entra redirect URIs → Re-upload manifests |
| Manifest or ribbon UI | `bun run manifests:generate --profile development` → Re-upload manifests |
| Production domain | Rebuild with production values → set `GE_PROD_WEB_DOMAIN` → `manifests:generate --profile production` + `package:prod` → add the domain's 3 redirect URIs in Entra → re-upload. Keep the `GE_PROD_*_APP_ID` values |
| Production apps (e.g. add OneNote) | Update `GE_PROD_SURFACES` (and add `GE_PROD_ONENOTE_APP_ID`) → regenerate → upload the new files |
| `skill/` files, **only if** `VITE_GE_COMMAND_PLANNER_SKILL` / `VITE_GE_SURFACE_COMMANDER_SKILL` are set | `bun run ge:skills` (uploads the bundles to Gemini Enterprise), then rebuild. Without those variables the app sends its own command grammar every turn and nothing needs uploading. A pane that shows "Model instructions out of date" in Activity needs this step |

---

## 9. Troubleshooting

| Issue | Likely Cause | Resolution |
| --- | --- | --- |
| Panel is blank | URL mismatch or build error | Verify web app origin matches manifest and Entra. Inspect task pane console (⌥+⌘+I). |
| Sign-in fails in Outlook | Missing `brk-` redirect links in Entra | Add `brk-multihub://` and `brk-9199bf20-...` to Entra SPA redirect URIs. |
| Button not visible in Outlook | Looking at inbox folder | Open an individual email or compose window; the add-in button lives on the message surface. |
| Add-in missing from Outlook's add-in list after deploying | It was deployed or uploaded, but not added to your Outlook yet (admin deployments can take up to 24 hours) | Open [aka.ms/olksideload](https://aka.ms/olksideload) → **My add-ins** → scroll to the bottom → **Add a custom add-in** → **Add from file...** → `dist/package/development/centralized/outlook.manifest.xml`. |
| Manifest validation fails | Invalid XML or missing icon | Run `bun run manifests:validate --profile development` to inspect errors before uploading. |
| Old code displays after deploy | Client browser cache | Close task pane, clear Office cache, or reload the frame. |
| `Missing production manifest configuration: …` | A `GE_PROD_*` value isn't set | Add the named value to `packages/web-shell/.env` or the CI environment (section 6). |
| `GE_PROD_*_APP_ID must be a non-development GUID` / `must all be different` | A placeholder or reused app ID | Generate a new ID with `uuidgen \| tr A-Z a-z` for that value only. Don't change IDs that are already in use. |
| Two copies of ContosoAI in Office | An app ID changed between uploads | Put the original ID back in `.env`, regenerate, re-upload, and remove the extra copy in the admin center. |
