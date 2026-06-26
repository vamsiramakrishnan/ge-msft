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

Example CLI shape:

```bash
npx @microsoft/teamsapp-cli install --file-path dist/release/development-m365-v<version>.zip
```

If the CLI opens a browser in a headless workstation and fails with `xdg-open ENOENT`, copy the URL
into your browser manually or use a device-code/non-browser login flow if available.

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
