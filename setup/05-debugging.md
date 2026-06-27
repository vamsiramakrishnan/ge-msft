# Debugging and Troubleshooting

## Unsupported Host in a Popup

`Unsupported host` in a sign-in popup usually means the popup loaded `taskpane.html` outside Office.
The auth redirect should load:

```text
https://<origin>/auth-redirect.html
```

not the task pane route. Verify the Entra SPA redirect URI and `VITE_ENTRA_CLIENT_ID`.

## MSAL Popup Timeouts

Common causes:

- `auth-redirect.html` is not registered as an Entra SPA redirect URI.
- The redirect host changed after regenerating/uploading the manifest.
- Browser blocked the popup or an old popup is still open.
- The add-in iframe is sandboxed and top-level redirect was attempted instead of popup/redirect
  bridge behavior.

Recovery:

1. Close all Microsoft sign-in popups.
2. Sign out/in of the Office web app if needed.
3. Clear site data for the add-in origin only.
4. Reopen the task pane and click Sign in once.

## Cloudflare 502 or Refused to Connect

Check:

```bash
curl -k https://localhost:13000/taskpane.html
```

If local curl fails, restart Vite. If local curl works but the tunnel fails, restart `cloudflared`
and re-upload the regenerated manifest with the new tunnel hostname.

## WIF invalid_grant Audience Mismatch

Error shape:

```text
invalid_grant: The audience in ID Token [...] does not match the expected audience.
```

Ask the Google Cloud admin to add/trust the Entra application client ID used by:

```text
VITE_ENTRA_CLIENT_ID
```

on the Workforce Identity provider. The add-in and the provider must agree on the same Entra app
client ID.

## Gemini Catalog Permission Errors

If `discoveryengine.googleapis.com/.../agents` returns 403, use the widget endpoints if your Gemini
Enterprise app exposes them:

```text
widgetListAvailableAgentViews
lookupWidgetConfig
```

Set:

```bash
VITE_GE_WIDGET_CONFIG_ID=<widget-config-id>
VITE_GE_WIDGET_SERVER_TOKEN=<optional-server-token>
```

If the signed-in user still cannot list skills, configure fallback skill resources in `.env`.

## Excel /visualize Produces Prose Instead of a Chart

The command loop must mount the command skill and receive exactly one fenced `cmd` block. The fixed
runtime path sends command-loop turns with the command skill route and rejects `done` batched after a
write. Correct model output shape:

```cmd
read 'Project schedule'!B5:F30
```

then after the result:

```cmd
chart bar 'Project schedule'!B5:D30 title="Task Progress"
```

The runtime previews the chart insertion and asks for approval before Excel is mutated.
