// ADR-0008 §4 — emit the versioned language manifest from the authoritative @ge/contracts
// definitions into a committed JSON the skill's Python preflight loads. Run after building
// contracts (`npm run typecheck` / `tsc -b` produces dist/). Re-running must produce NO diff —
// a byte change is a real language change (CI diffs the committed file as the drift gate).
import { writeFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildLanguageManifest, assertManifestConsistent } from '../dist/language-manifest.js';

const manifest = assertManifestConsistent(buildLanguageManifest());
const json = JSON.stringify(manifest, null, 2) + '\n';

const here = dirname(fileURLToPath(import.meta.url));
const targets = [
  // Bundled into the skill package so the Python sandbox loads it (no hand-maintained tables).
  resolve(here, '../../../skill/m365-surface-commander/scripts/m365-cli-1.0.json'),
];
for (const out of targets) {
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, json);
  console.log(`wrote ${out} (${manifest.version}, ${manifest.verbs.write.length} write verbs)`);
}
