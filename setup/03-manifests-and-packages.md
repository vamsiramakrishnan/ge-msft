# Manifest Generation and Packages

There are three different development artifacts. Use the one that matches the host/client you are
testing.

## Generate Manifests

Development profile:

```bash
bun run manifests:generate -- --profile development
bun run manifests:validate -- --profile development
```

Outputs:

```text
dist/manifests/development.manifest.json
dist/manifests/development.word.manifest.xml
dist/manifests/development.excel.manifest.xml
dist/manifests/development.powerpoint.manifest.xml
dist/manifests/development.outlook.manifest.xml
dist/manifests/development.onenote.manifest.xml
```

Internal alpha profile:

```bash
bun run manifests:generate -- --profile internal-alpha-word-excel
bun run manifests:validate -- --profile internal-alpha-word-excel
```

The alpha profile intentionally advertises Word and Excel only.

## Build and Package Development

```bash
bun run build
bun run package:dev
```

Outputs:

```text
dist/release/development-m365-v<version>.zip
dist/package/development/m365/manifest.json
dist/package/development/xml/word.manifest.xml
dist/package/development/xml/excel.manifest.xml
dist/package/development/xml/powerpoint.manifest.xml
dist/package/development/xml/outlook.manifest.xml
dist/package/development/onenote/onenote.manifest.xml
dist/release/development-artifact.json
dist/release/SHA256SUMS
```

## Which Artifact Is Which

- Unified M365 package zip:
  - `dist/release/development-m365-v<version>.zip`
  - Contains root `manifest.json` and icons.
  - Use with Teams/M365 app package upload flows and Agents Toolkit.
- Classic Office XML manifests:
  - `dist/package/development/xml/*.manifest.xml`
  - Use when Word/Excel/PowerPoint/Outlook "Upload Add-in" asks for a manifest file.
  - The upload dialog expects XML, not the unified zip.
- OneNote XML manifest:
  - `dist/package/development/onenote/onenote.manifest.xml`
  - OneNote uses a separate legacy XML flow.

## Package Alpha

```bash
bun run build
bun run manifests:generate -- --profile internal-alpha-word-excel
bun run manifests:validate -- --profile internal-alpha-word-excel
bun run package:alpha
```

Output:

```text
dist/release/internal-alpha-word-excel-v<version>.zip
dist/release/artifact.json
dist/release/SHA256SUMS
```

Do not use alpha packaging for all-surface development. It is intentionally Word + Excel only.
