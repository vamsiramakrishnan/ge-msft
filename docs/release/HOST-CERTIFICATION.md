# Live-Host Certification Protocol

"Live-host certification" is the manual, per-surface verification of a **specific built package**
on **real Office hosts** — not simulators, not the unit-test fakes — before an alpha or production
cut ships. The automated suite proves the bridges against faked host object models; this protocol
proves that the same bytes boot, authenticate, capture, write, and record provenance inside actual
Word and Excel, on desktop and on the web.

The release gate (`tools/release/release-check.mjs`, `certificationOutcomes()`) enforces it: the
gate stays **BLOCKED** for each required surface until a matching evidence report exists, and
**FAILS** if a report exists but does not match the exact release under check.

## What the gate checks (match this exactly)

- **Location:** JSON files under `evidence/host-certification/` (any `*.json` in that directory
  tree is considered; see `evidence/host-certification/README.md` for the naming convention).
- **Required surfaces:** `word` and `excel` (the gate iterates `['word', 'excel']`).
- **Matching:** a report counts for a surface when `data.surface === '<surface>'` **and**
  `data.profile === <the profile being checked>` (e.g. `internal-alpha-word-excel`).
- **Pinning:** the matched report must satisfy, against the release under check:
  - `commitSha` — equals `git rev-parse HEAD` of the release tree,
  - `packageHash` — equals the SHA-256 of the package zip (`dist/release/<profile>-v<version>.zip`;
    printed by the packager and recorded in the artifact json),
  - `manifestVersion` — equals the `version` field of the generated manifest,
  - every entry in `tests[]` has `status: "pass"` (any other status fails the gate).

Report shape:

```json
{
  "surface": "word",
  "profile": "internal-alpha-word-excel",
  "commitSha": "<git rev-parse HEAD>",
  "packageHash": "<sha256 of the package zip>",
  "manifestVersion": "0.1.0",
  "host": { "app": "Word", "platform": "desktop-windows", "build": "16.0.xxxxx" },
  "certifiedBy": "<name/alias>",
  "certifiedAt": "2026-07-06T12:00:00Z",
  "tests": [
    { "name": "boot", "status": "pass", "notes": "pane loads at production origin" },
    { "name": "auth", "status": "pass", "notes": "NAA sign-in -> WIF exchange -> assist ok" },
    { "name": "capture", "status": "pass" },
    { "name": "gated-write", "status": "pass" },
    { "name": "provenance", "status": "pass" }
  ]
}
```

`host`, `certifiedBy`, `certifiedAt`, and `notes` are not read by the gate but are required by
this protocol for the audit trail. One file per surface × platform run is fine; the gate needs at
least one fully passing report per surface for the exact commit + package + manifest version.

**Certification never survives a rebuild.** Any new commit or repackage changes `commitSha` /
`packageHash`, so the evidence must be re-produced. That is intentional — do not edit an old
report's hashes to match a new build (see the README in the evidence directory).

## Platforms

Certify each required surface on **both** desktop (Windows or Mac, current channel) and the web
host (office.com), sideloading the exact package under certification (see `setup/04-sideloading.md`
for mechanics). Record each platform run in the report(s).

## Per-surface checklist

Run every item; every item maps to a `tests[]` entry. A skipped item means the report must not be
written.

### All surfaces

1. **boot** — Ribbon shows the Gemini Enterprise group; the task pane opens from the ribbon
   button; no manifest warning from the host; pane assets load from the release origin (verify in
   devtools/network — no localhost, no dev tunnel).
2. **auth** — Sign-in completes via NAA in the pane (no popup loop); the shell obtains the Entra
   token and completes the WIF exchange; a grounded `streamAssist` round trip returns; sign-out
   and re-sign-in works.
3. **capture** — The pane captures real document content (not the fake): verify the outline/
   selection context shown in the pane matches the open document.
4. **gated-write** — An agent-proposed change renders as a reviewable proposal first; nothing
   touches the document until explicitly applied; Cancel leaves the document byte-stable.
5. **provenance** — After applying, the provenance record (agent id, sources, identity, timestamp,
   content hash) is present in the host's durable metadata and survives save → close → reopen.

### Word specifics

- capture: outline + selection capture on a multi-page styled document (headings, tables).
- gated-write: applied change lands as a **tracked change** attributed correctly; Reject Change
  reverts it cleanly; a stale finding (edit the anchor text before applying) degrades to a panel
  item instead of writing at a wrong anchor (`body.search` re-resolution).
- provenance: record lands in the document's custom XML part
  (`packages/bridge-word/src/provenance-record.ts` path) and persists through save/reopen.

### Excel specifics

- capture: range/used-range capture on a workbook with multiple sheets and formats.
- gated-write: `write-cells` applies to the exact reviewed address; formats survive; the on-sheet
  result matches the preview; undo restores the prior values.
- provenance: record lands in the workbook settings path
  (`packages/bridge-excel/src/provenance-record.ts`) and persists through save/reopen.

### Production profile additions

When certifying `production` (all-surface package), run the All-surfaces checklist on PowerPoint,
Outlook (read + compose + on-send gate), OneNote (web, XML package), and Teams as well; keep
`word` and `excel` reports as the gate-required minimum and add the other surfaces as extra
reports with the same shape.

## Producing the report

```bash
git rev-parse HEAD                          # -> commitSha (tree must be the release commit, clean)
npm run manifests:generate -- --profile <profile>
npm run package:alpha                       # or: npm run package:prod
sha256sum dist/release/<profile>-v<ver>.zip # -> packageHash (also in dist/release/*artifact.json)
```

Sideload that exact zip, execute the checklist, then write the JSON report(s) into
`evidence/host-certification/` and re-run `npm run release:check` — the surface's gate should flip
from BLOCKED to pass.
