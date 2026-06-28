# Claude Office Add-in Observations

Date: 2026-06-27

Scope: clean-room inspection of publicly served artifacts from `https://pivot.claude.ai`. The
artifacts were downloaded to `/tmp/pivot-research` for analysis only. No Claude bundle, chunk, CSS,
or manifest content is vendored into this repository.

## Pulled Artifacts

The public surface exposes a classic Office XML manifest and a taskpane app shell:

| Artifact | Local analysis copy | SHA-256 |
| --- | --- | --- |
| Office XML manifest | `/tmp/pivot-research/manifest.xml` | `a4d276bf60ecdc38a32e4f7ab83b266449c3e5fdb49a54a3a09158b0c41f3a08` |
| Same manifest under `/m-addin` | `/tmp/pivot-research/m-addin-manifest.xml` | same content as above |
| Keyboard shortcuts JSON | `/tmp/pivot-research/shortcuts.json` | `f26c0d451df00417699000253622f3c8dc264993327975a3f20ca6161c98c86c` |
| App shell HTML | `/tmp/pivot-research/index.html` | `5d2a9aaee6d36a788b0d7b67b7c758bcb0a346bf5a0be946df268009c6a8800f` |
| Main app bundle | `/tmp/pivot-research/app.bundle.js` | `e31524881687ae793699061d865d8c25041e5f8dfd85308dd3b520d61578974e` |
| Main CSS | `/tmp/pivot-research/styles.css` | `87d97551db634e9c00d69fb99ee8634d658cb945e7825814071077269ba0aae3` |

The app bundle references 55 lazy chunks. The notable chunk names are safe to record because they are
asset filenames, not implementation source: `cat`, `file-reader`, `ls`, `grep`, `rg`, `jq`, `diff`,
`find`, `tree`, `wc`, `sort`, `uniq`, `cut`, `paste`, `xargs`, `base64`, `sha256sum`, `file`, and
related parsing/support chunks.

## Manifest Findings

The Office path is a classic XML manifest, not a Teams/unified manifest package. `/manifest.xml` is
served directly as XML, while `/manifest.json` falls back to the SPA HTML shell.

Key structure observed:

- One XML package covers Word, Excel, and PowerPoint by declaring `Document`, `Workbook`, and
  `Presentation` hosts.
- The manifest uses `SharedRuntime` and a long-lived taskpane runtime for each host.
- Ribbon entry is declared through `VersionOverrides`, `DesktopFormFactor`, `PrimaryCommandSurface`,
  and host-specific button IDs.
- The taskpane URL is a single app shell URL with a version-like query parameter, instead of separate
  per-host `taskpane.html?host=...` URLs.
- Keyboard shortcut support is declared through an `ExtendedOverrides` URL pointing at a separate
  JSON file.
- The public shortcut file declares one taskpane action and one default keyboard shortcut.
- Outlook is not included in this XML shape; an Outlook add-in still needs a separate `MailApp`
  manifest if we want mail compose/read behavior.
- OneNote remains separate for our architecture because it uses a different legacy web-only path.

## Bundle Capability Signals

The bundle appears to include a browser-local file/workspace analysis layer:

- It lazy-loads small Unix-like tool chunks for file inspection, search, tabular transforms, checksums,
  and diffs.
- Term scans show heavy usage of host/document concepts across Excel, Word, PowerPoint, Google Sheets,
  Google Docs, and Google Slides.
- CSS includes a dedicated response typography system, editor styling, dark-mode tokens, and compact
  button states.
- The bundle includes Sentry release metadata and MSAL code, suggesting production diagnostics and
  Microsoft identity bootstrap are part of the deployed shell.

This does not justify copying the implementation. It does validate a product pattern worth building
clean-room: a constrained in-browser analysis workspace that can inspect user-approved snapshots
before deciding whether to call the model, upload context files, search connectors, or request host
reads.

## Local Gaps Exposed

### 1. Development XML Manifests Are Too Thin

Our generated Word/Excel/PowerPoint XML manifests are basic taskpane manifests. The generator at
`tools/release/common.mjs` currently emits host-specific XML in `taskPaneXmlManifest(...)` with base
`DefaultSettings` only. It lacks the richer `VersionOverrides`, `bt` resources, host-specific ribbon
commands, shared runtime, function file, and optional keyboard shortcut extension used by the public
Claude XML.

Local fix point:

- `tools/release/common.mjs`, `taskPaneXmlManifest(...)`

Recommended correction:

- Add a second development XML artifact that covers Word + Excel + PowerPoint in one classic
  `TaskPaneApp` manifest.
- Keep existing host-specific XML files for Office web upload flows that reject a multi-host manifest.
- Add `VersionOverridesV1_0` with per-host `DesktopFormFactor`, `PrimaryCommandSurface`, icons, URLs,
  strings, and a `FunctionFile` pointing at `commands.html`.
- Add an optional generated `shortcuts.json` and `ExtendedOverrides` only when the command runtime is
  registered and tested.
- Keep Outlook as its existing `MailApp` XML.

### 2. Browser Analysis Workspace Is Missing

We already have `ContextFileClient.addContextFile(...)`, guarded file normalization, and StreamAssist
`fileIds` support. The missing layer is an explicit browser-local workspace that lets the model and UI
choose among local analysis, host reads, connector search, and upload.

Local fix points:

- `packages/gemini-client/src/context-files.ts`
- `packages/gemini-client/src/stream-assist.ts`
- `packages/web-shell/src/taskpane/components/App.tsx`
- `packages/web-shell/src/taskpane/components/GeminiCatalogPanel.tsx`

Recommended correction:

- Introduce a `@ge/browser-workspace` or `packages/web-shell/src/workspace` module with memory-only
  files, byte caps, MIME allow-listing, content hashes, and explicit lifecycle.
- Add safe, allow-listed tools: `head`, `tail`, `wc`, `rg`, `jq`, `csv-profile`, `diff`, `sha256`, and
  workbook/document outline extractors.
- Sanctioned analytical engine (deliberate deviation, recorded here): the allow-list above can
  profile/search/diff but cannot do real tabular analytics — multi-table joins, group-by, window
  functions, pivots. A pure safe-tool set hits a **capability wall** for the Excel analysis use case.
  Include **DuckDB-WASM as a single, hardened, read-only analytical tool**, exposed through ONE typed
  `query` command (not a shell). It must be locked down: extensions disabled (no `INSTALL`/`LOAD`), no
  `httpfs`/network, no host filesystem access (`ATTACH`/`COPY`-to-disk/`EXPORT` blocked); it operates
  only over tables the workspace explicitly seeds from user-approved snapshots/uploads (Arrow/CSV held
  in memory). Queries are read-only; results return as data, and any write-back to the host still goes
  through the existing dry-run → preview → approval → actuation gate.
- Still banned: general shell execution, `eval`, arbitrary JS, arbitrary Office Scripts, PowerShell,
  and process execution. The DuckDB `query` command is the ONLY code-bearing surface, and it is
  sandboxed, read-only, and offline by construction (no-network Web Worker, no host/DOM access).
- Model-facing access should be through constrained, typed commands such as `context`, `inspect`,
  `attach`, and `query` — never a general shell.

### 3. Context Picker Is Config-Oriented, Not Work-Oriented

Our catalog UI lists skills and connectors as routing config. Users need a task-level context picker:
current selection, current document snapshot, connector search results, recent files, and uploaded
files.

Local fix points:

- `packages/web-shell/src/taskpane/components/GeminiCatalogPanel.tsx`
- `packages/web-shell/src/taskpane/components/ContextTray.tsx`

Recommended correction:

- Keep skill routing behind settings.
- Promote connector nodes/search/upload/current-file attachment into the context tray.
- Show attached context as compact chips with hover detail; open a picker only on demand.
- Persist only non-secret metadata and hashes; never persist content or tokens.

### 4. Response Rendering Still Lags

Our renderer is intentionally safe and now handles basic markdown, but the inspected CSS indicates a
more complete response/editor system in the public Claude shell.

Local fix point:

- `packages/web-shell/src/taskpane/components/MessageThread.tsx`

Recommended correction:

- Keep URL sanitization.
- Add copy controls, collapsible tool/code-execution traces, table overflow handling, source hover
  cards, and compact provenance drill-in.
- Do not render untrusted HTML.

### 5. Product Surface Should Be Quieter

The current pane exposes a lot of persistent operational chrome: catalog, skills, context, command
center, quick actions, run steps, gates, and composer. The public artifact pattern points toward a
single app shell with richer progressive disclosure.

Local fix points:

- `packages/web-shell/src/taskpane/components/App.tsx`
- `packages/web-shell/src/taskpane/components/SurfaceCommandCenter.tsx`
- `packages/web-shell/src/taskpane/components/QuickActionBar.tsx`
- `packages/web-shell/src/taskpane/styles.css`

Recommended correction:

- Default pane: header, compact active context, top 2-3 contextual actions, thread, composer.
- Move catalog, skills, detailed run steps, and extended actions behind hover/click disclosure.
- Keep approval cards prominent only when there is a live decision gate.

## Clean-Room Build Plan

1. Manifest parity:
   - Generate a multi-host classic XML manifest for Word + Excel + PowerPoint.
   - Add resources, ribbon controls, shared runtime, command function file, and optional shortcut
     extension.
   - Validate with existing manifest tooling plus an Office XML validator if available.

2. Browser workspace:
   - Add a memory-only VFS abstraction with hashes and byte/type limits.
   - Implement a safe analysis tool registry, plus the hardened read-only DuckDB-WASM `query` tool
     (extensions/network/filesystem disabled; seeded in-memory tables only).
   - Run all compute in a no-network Web Worker with no host/DOM access; enforce per-query timeouts and
     memory caps.
   - Add tests proving: no general shell/process execution; DuckDB cannot install/load extensions,
     reach the network, or read/write host files; queries are read-only; and no durable token/content
     persistence.

3. Planner/commander integration:
   - Teach the planner to choose `incremental`, `local-analysis`, `upload-preferred`, or
     `connector-search` as an explicit strategy.
   - Teach the commander to request workspace inspection through typed commands only.

4. Context UX:
   - Replace connector checkboxes as the primary user affordance with a contextual source picker.
   - Keep admin-style skill routing in settings.

5. Response UX:
   - Upgrade markdown/table/source rendering while preserving URL sanitization and content distrust.

## Non-Goals

- Do not copy, de-minify into source, or vendor the public Claude app bundle.
- Do not add arbitrary shell/process execution to the browser.
- Do not expose a general shell, `eval`, or arbitrary code execution. The ONLY code-bearing surface is
  the sandboxed, read-only, offline DuckDB `query` tool over user-approved in-memory snapshots.
- Do not weaken the existing parse -> validate -> dry-run -> preview -> approval -> actuation path.
- Do not ship copied branding, icons, text, or proprietary UI implementation details.
