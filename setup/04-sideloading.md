# Sideloading

Microsoft has more than one sideloading path. The most common confusion is zip versus XML.

## Word, Excel, PowerPoint, Outlook - Office Upload Add-in

When Office shows:

```text
Upload Add-in
Choose your add-in manifest
```

upload the XML manifest for that host, not the zip:

```text
dist/package/development/xml/word.manifest.xml
dist/package/development/xml/excel.manifest.xml
dist/package/development/xml/powerpoint.manifest.xml
dist/package/development/xml/outlook.manifest.xml
```

Typical Office on the web path:

1. Open Word, Excel, PowerPoint, or Outlook on the web.
2. Open Add-ins.
3. Choose Upload Add-in.
4. Select the host-specific XML manifest.
5. Open the Gemini Enterprise command from the ribbon.

If the upload dialog rejects the zip, that is expected. Use XML.

## Unified Manifest / M365 Package Zip

For the full deployment-method distinction, including app catalog versus Office add-in deployment,
see [Deployment methods matrix](./07-deployment-methods-matrix.md).

Use the zip for the unified Microsoft 365 app package flow:

```text
dist/release/development-m365-v<version>.zip
```

The zip contains:

```text
manifest.json
icon-color.png
icon-outline.png
other icon assets
```

Possible upload paths:

- Microsoft 365 / Teams developer app upload, when custom app upload is allowed by the tenant.
- Microsoft 365 Agents Toolkit / Teams Toolkit CLI.

### Automated Developer Sideload With Agents Toolkit

Use this when you want a local developer install of the unified package without manually uploading
the zip in the Teams/M365 UI.

```bash
bun run sideload
```

That command:

1. ensures `cloudflared` is available,
2. starts/restarts Vite and a Cloudflare quick tunnel,
3. writes the new tunnel origin to `packages/web-shell/.env`,
4. regenerates, validates, and packages the development manifests using that origin,
5. patches the Entra SPA redirect to include `https://<tunnel>/auth-redirect.html`,
6. checks Agents Toolkit Microsoft 365 auth and prompts for `atk auth login m365` when needed,
7. runs `atk install --file-path <zip>`,
8. records the returned Agents Toolkit title ID in `.ge-sideload/unified-sideload.json` when it can
   parse it.

Useful variants:

```bash
bun run sideload                       # preflight ATK auth, then install/update
bun run sideload -- --login            # force a fresh atk auth login m365 first
bun run sideload -- --skip-atk-login   # skip ATK auth preflight
bun run sideload -- --skip-tunnel      # package/install only; use when the tunnel is already correct
bun run sideload:status                # show remembered title ID/package
bun run sideload:uninstall             # uninstall remembered title ID
bun run m365:sideload                  # interactive install/uninstall/status flow
bun run setup:atk:login                # only run/check atk auth login m365
```

Equivalent raw CLI shape:

```bash
atk install --file-path dist/release/development-m365-v<version>.zip
atk uninstall --mode title-id --title-id U_<title-id-guid> --interactive false
```

Microsoft's current unified-manifest sideload guidance says the Agents Toolkit CLI returns a title
ID and that uninstall should use that title ID. If our wrapper cannot parse the title ID from `atk`
output, copy it from the terminal and run:

```bash
bun run sideload:uninstall -- --title-id U_<title-id-guid>
```

If the CLI opens a browser in a headless workstation and fails with `xdg-open ENOENT`, copy the URL
into your browser manually or run the scripted login helper:

```bash
bun run setup:atk:login
```

That helper wraps `atk auth login m365` and reuses the existing account when `atk auth list` already
shows a connected Microsoft 365 account. Then retry `bun run sideload`.

This is still a **developer sideload**. It is not tenant rollout, assignment, pinning, or guaranteed
availability across every Office client. For tenant catalog upload, use:

```bash
bun run m365:catalog
```

## OneNote

OneNote uses the separate legacy manifest:

```text
dist/package/development/onenote/onenote.manifest.xml
```

Use the OneNote web or admin-center/shared-catalog flow available in your tenant. Do not upload the
unified zip to the OneNote XML flow.

## Desktop Office

Desktop Office can sideload classic Office XML manifests through the platform-specific trusted
catalog or shared-folder mechanism. Use the host-specific XML manifest from:

```text
dist/package/development/xml/
```

Keep the dev server or deployed web origin running. The desktop host still loads the task pane from
the manifest URL.

## After Upload

If you do not see the add-in:

- Confirm the manifest origin matches the running server or tunnel.
- Confirm the tenant allows custom add-ins/custom app upload.
- Refresh the Office web page or restart desktop Office.
- Re-upload after changing tunnel hostnames.
- Check that the manifest type matches the upload UI: XML for Office Upload Add-in, zip for unified
  M365 app package upload.

## References

- Microsoft unified manifest sideloading:
  <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest>
- Microsoft Office on the web add-in sideloading:
  <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing>
- Microsoft Outlook add-in sideloading:
  <https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing>
