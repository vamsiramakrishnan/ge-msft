# Prerequisites and Configuration

## What You Need

- A Microsoft 365 tenant with access to Microsoft Entra app registrations.
- A Google Cloud / Gemini Enterprise app that the signed-in Microsoft user can access through
  Workforce Identity Federation.
- Bun 1.3 or newer. Node 20 or newer is still required by some workspace tooling.
- Optional: Azure CLI (`az`) or Microsoft 365 Agents Toolkit CLI for app registration and package
  upload helpers. An Azure subscription is not required just to create an Entra app registration in
  a Microsoft 365 tenant, but your account must have permission to create or edit app registrations.
- Optional for remote workstations: `cloudflared` for a public HTTPS dev tunnel.

## Local Web-Shell Config

Copy:

```bash
cp packages/web-shell/.env.example packages/web-shell/.env
```

Fill only public client configuration. Do not add client secrets, Google service-account keys,
private keys, refresh tokens, or durable bearer tokens.

Required local values:

```bash
VITE_GCP_PROJECT=<gcp-project-id-or-number>
VITE_GCP_LOCATION=global
VITE_GE_ENGINE=<gemini-enterprise-engine-id>
VITE_GE_COLLECTION=default_collection
VITE_GE_ASSISTANT=default_assistant

VITE_WIF_POOL_ID=<workforce-pool-id>
VITE_WIF_PROVIDER_ID=<provider-id>

VITE_ENTRA_CLIENT_ID=<entra-public-client-app-id>
VITE_ENTRA_AUTHORITY=https://login.microsoftonline.com/<tenant-id-or-domain>
VITE_WIF_ID_TOKEN_SCOPES=openid profile email User.Read
VITE_GRAPH_SCOPES=User.Read

VITE_GE_WIDGET_CONFIG_ID=<optional-widget-config-id>
VITE_GE_WIDGET_SERVER_TOKEN=<optional-widget-server-token>
```

If the signed-in user cannot list Gemini Enterprise skills/connectors, set fallback skill resources:

```bash
VITE_GE_COMMAND_PLANNER_SKILL=m365-command-planner=<full-agent-resource>
VITE_GE_SURFACE_COMMANDER_SKILL=m365-surface-commander=<full-agent-resource>
```

The task pane can also discover these through the Gemini Enterprise catalog dropdown when the user
has permission.

## Microsoft Entra App Registration

Create or reuse a public-client app registration for the add-in.

Required fields:

- Tenant ID or primary tenant domain.
- Application client ID.
- Single-page application redirect URI:
  - `https://<dev-origin>/auth-redirect.html`
  - `https://<production-origin>/auth-redirect.html` for a deployed environment.
- API permissions:
  - `openid`
  - `profile`
  - `email`
  - `User.Read`
  - Add only narrow delegated Graph scopes when a feature needs them.
- No client secret. This is a browser public client.

If you use a Cloudflare tunnel for development, the tunnel hostname is part of the redirect URI and
must be added to this app registration.

## Google Cloud / WIF Values

Ask the Google Cloud administrator for:

- GCP project ID or number.
- Gemini Enterprise location, usually `global`.
- Gemini Enterprise engine/application ID.
- Workforce pool ID.
- Workforce provider ID.
- Expected Entra audience/client ID configured on the provider.
- Optional widget config ID and server token for `widgetListAvailableAgentViews` and
  `lookupWidgetConfig`.

The WIF provider must trust the same Entra app client ID used by `VITE_ENTRA_CLIENT_ID`. If Google
STS returns `invalid_grant` with an audience mismatch, the Entra client ID in the ID token does not
match the provider's allowed audience.

## Useful Login Checks

Google CLI login through Workforce Identity Federation must use the Microsoft SSO path configured by
the Google admin. If `gcloud auth login` sends you to a consumer Google account, use the workforce
provider login command or browser link provided by the Google admin instead of a normal Google login.

Azure CLI, if available:

```bash
az login --tenant <tenant-id-or-domain> --use-device-code
az account show
```

If browser redirect to localhost fails in a remote workstation, use device code login when the CLI
supports it.
