#!/usr/bin/env node
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  cleanDir,
  copyDir,
  copyFile,
  createZip,
  generatedManifestPath,
  packageDir,
  packageZip,
  parseArgs,
  profileFromArgs,
  repoRoot,
  rootVersion,
  sha256File,
  walk,
  writeChecksums,
  writeJson,
} from './common.mjs';

const args = parseArgs();
const profile = profileFromArgs(args);
if (profile !== 'internal-alpha-word-excel') {
  console.error('package:alpha only packages internal-alpha-word-excel.');
  process.exit(1);
}

const manifest = generatedManifestPath(profile);
const web = join(repoRoot, 'packages', 'web-shell', 'dist-web');
if (!existsSync(manifest)) {
  console.error(`Generated alpha manifest missing: ${manifest}`);
  process.exit(1);
}
if (!existsSync(web)) {
  console.error(`Built web shell missing: ${web}. Run npm run build first.`);
  process.exit(1);
}

const outDir = packageDir(profile);
cleanDir(outDir);
copyDir(web, outDir);
copyFile(manifest, join(outDir, 'manifest.json'));

// The unified manifest names a stable commands script. Vite emits a hashed chunk; copy the first
// command entry chunk to the stable package path while keeping the original HTML/chunks intact.
const commandsChunk = walk(join(outDir, 'assets')).find((file) => /commands-.*\.js$/.test(file));
if (commandsChunk) copyFile(commandsChunk, join(outDir, 'assets', 'commands.js'));

const releaseNotes = [
  `# Gemini Enterprise Internal Alpha v${rootVersion()}`,
  '',
  'Profile: internal-alpha-word-excel',
  'Enabled surfaces: Word, Excel',
  'Disabled surfaces: PowerPoint, OneNote, Outlook, Teams',
  '',
  'This package is not an AppSource/public beta artifact.',
  '',
].join('\n');
writeFileSync(join(outDir, 'RELEASE-NOTES.md'), releaseNotes);

const zipPath = packageZip(profile);
createZip(walk(outDir), outDir, zipPath);
const checksumPath = join(repoRoot, 'dist', 'release', 'SHA256SUMS');
writeChecksums([zipPath, manifest], checksumPath);
writeJson(join(repoRoot, 'dist', 'release', 'artifact.json'), {
  profile,
  version: rootVersion(),
  package: zipPath,
  packageSha256: sha256File(zipPath),
  manifest,
  manifestSha256: sha256File(manifest),
  manifestVersion: JSON.parse(readFileSync(manifest, 'utf8')).version,
});

console.log(`packaged ${zipPath}`);
console.log(`sha256 ${sha256File(zipPath)}`);
