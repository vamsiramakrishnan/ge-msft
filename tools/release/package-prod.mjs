#!/usr/bin/env node
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  artifactPlaceholderFindings,
  cleanDir,
  copyDir,
  copyFile,
  createZip,
  generatedManifestPath,
  generatedOneNoteManifestPath,
  isPlaceholderGuid,
  packageDir,
  parseArgs,
  profileFromArgs,
  repoRoot,
  rootVersion,
  scanForbiddenSecrets,
  sha256File,
  validateGeneratedManifest,
  walk,
  writeChecksums,
  writeJson,
} from './common.mjs';

const args = parseArgs();
const profile = profileFromArgs({ profile: 'production', ...args });
if (profile !== 'production') {
  console.error('package:prod only packages the production profile.');
  process.exit(1);
}

const manifest = generatedManifestPath(profile);
const oneNoteManifest = generatedOneNoteManifestPath(profile);
const web = join(repoRoot, 'packages', 'web-shell', 'dist-web');
const publicDir = join(repoRoot, 'packages', 'web-shell', 'public');
const requiredIcons = [
  'icon-color.png',
  'icon-outline.png',
  'icon-16.png',
  'icon-32.png',
  'icon-64.png',
  'icon-80.png',
];

if (!existsSync(manifest)) {
  console.error(
    `Generated production manifest missing: ${manifest}. Run npm run manifests:generate:prod first.`,
  );
  process.exit(1);
}
if (!existsSync(oneNoteManifest)) {
  console.error(`Generated production OneNote manifest missing: ${oneNoteManifest}`);
  process.exit(1);
}
if (!existsSync(web)) {
  console.error(`Built web shell missing: ${web}. Run npm run build first.`);
  process.exit(1);
}
for (const icon of requiredIcons) {
  const iconPath = join(publicDir, icon);
  if (!existsSync(iconPath)) {
    console.error(`Icon missing: ${iconPath}`);
    process.exit(1);
  }
}

// Validate the generated manifest again at package time so a stale dist/manifests file cannot
// slip into the production zip.
const manifestJson = JSON.parse(readFileSync(manifest, 'utf8'));
const manifestErrors = validateGeneratedManifest(manifestJson, profile);
if (manifestErrors.length) {
  for (const error of manifestErrors) console.error(error);
  process.exit(1);
}

const outDir = packageDir(profile);
const m365Dir = join(outDir, 'm365');
const oneNoteDir = join(outDir, 'onenote');
const webDir = join(outDir, 'web');
cleanDir(outDir);

copyFile(manifest, join(m365Dir, 'manifest.json'));
for (const icon of requiredIcons) copyFile(join(publicDir, icon), join(m365Dir, icon));
copyFile(oneNoteManifest, join(oneNoteDir, 'onenote.manifest.xml'));
copyDir(web, webDir);

// The unified manifest references the stable ${origin}/assets/commands.js path; Vite emits a
// hashed chunk. Copy it to the stable path — and fail loudly if the chunk is missing, because a
// package without it ships a manifest pointing at a 404 commands runtime.
const commandsChunk = walk(join(webDir, 'assets')).find((file) => /commands-.*\.js$/.test(file));
if (!commandsChunk) {
  console.error(
    `Commands entry chunk (assets/commands-*.js) not found under ${join(webDir, 'assets')}. ` +
      'The manifest references assets/commands.js; rebuild the web shell (npm run build).',
  );
  process.exit(1);
}
copyFile(commandsChunk, join(webDir, 'assets', 'commands.js'));

const releaseNotes = [
  `# Gemini Enterprise for Microsoft 365 v${rootVersion()}`,
  '',
  'Profile: production',
  'Unified package surfaces: Word, Excel, PowerPoint, Outlook, Teams',
  'Companion package: OneNote legacy XML manifest',
  '',
  'Deploy the web build to the production static origin, upload m365/ (zipped) through the',
  'Microsoft 365 admin center, and register onenote/onenote.manifest.xml separately.',
  'See docs/release/DEPLOYMENT.md and docs/release/ROLLBACK.md.',
  '',
].join('\n');
writeFileSync(join(outDir, 'README.md'), releaseNotes);

// Post-layout gates: no Google/Entra secret material, no placeholder identifiers or dev hosts
// anywhere in the artifact tree. The manifest dirs get the strict scan (localhost forbidden);
// the web bundle allows the literal "localhost" because the shell's own URL-validation code
// contains it as a string constant, but must not contain placeholder GUIDs, REPLACE_ markers,
// or example.com.
const secretFindings = scanForbiddenSecrets([outDir]);
if (secretFindings.length) {
  console.error(`Forbidden secret patterns found: ${JSON.stringify(secretFindings, null, 2)}`);
  process.exit(1);
}
const placeholderFindings = [
  ...artifactPlaceholderFindings([m365Dir, oneNoteDir]),
  ...artifactPlaceholderFindings([webDir], { allowLocalhost: true }),
];
if (placeholderFindings.length) {
  console.error(`Placeholder/dev-host content found in: ${placeholderFindings.join(', ')}`);
  process.exit(1);
}
const GUID_ANYWHERE = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi;
const guidFindings = [];
for (const file of walk(webDir)) {
  if (!/\.(js|css|html|json)$/i.test(file)) continue;
  for (const guid of readFileSync(file, 'utf8').match(GUID_ANYWHERE) ?? []) {
    if (isPlaceholderGuid(guid)) guidFindings.push(`${file}: ${guid}`);
  }
}
if (guidFindings.length) {
  console.error(`Development placeholder GUIDs found in web bundle: ${guidFindings.join(', ')}`);
  process.exit(1);
}

const zipPath = join(repoRoot, 'dist', 'release', `production-m365-v${rootVersion()}.zip`);
createZip(walk(m365Dir), m365Dir, zipPath);

const checksumPath = join(repoRoot, 'dist', 'release', 'SHA256SUMS');
writeChecksums([zipPath, manifest, oneNoteManifest], checksumPath);

writeJson(join(repoRoot, 'dist', 'release', 'production-artifact.json'), {
  profile,
  version: rootVersion(),
  m365Package: zipPath,
  m365PackageSha256: sha256File(zipPath),
  m365Manifest: manifest,
  m365ManifestSha256: sha256File(manifest),
  oneNoteManifest,
  oneNoteManifestSha256: sha256File(oneNoteManifest),
  manifestVersion: manifestJson.version,
  webBuild: webDir,
});

console.log(`packaged ${zipPath}`);
console.log(`sha256 ${sha256File(zipPath)}`);
console.log(`onenote ${join(oneNoteDir, 'onenote.manifest.xml')}`);
console.log(`web ${webDir}`);
