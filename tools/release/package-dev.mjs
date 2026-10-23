#!/usr/bin/env node
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { BRAND_ICONS } from '../brand/brand.mjs';
import {
  BRAND,
  alphaProfile,
  prodProfile,
  cleanDir,
  copyDir,
  copyFile,
  createZip,
  generatedManifestPath,
  generatedOneNoteManifestPath,
  generatedOfficeXmlManifestPath,
  packageDir,
  parseArgs,
  profileFromArgs,
  releaseConfig,
  repoRoot,
  rootVersion,
  sha256File,
  walk,
  writeChecksums,
  writeJson,
} from './common.mjs';

// Packages the development profile (default) or the production profile.
const profile = profileFromArgs({ profile: 'development', ...parseArgs() });
if (profile === alphaProfile) {
  console.error('Use bun run package:alpha for the alpha profile.');
  process.exit(1);
}
const isDev = profile !== prodProfile;
const surfaces = isDev
  ? ['word', 'excel', 'powerpoint', 'outlook', 'onenote']
  : releaseConfig(profile).surfaces;
const hasDocumentHost = surfaces.some((s) => ['word', 'excel', 'powerpoint'].includes(s));
const label = isDev ? 'development' : 'production';
const manifest = generatedManifestPath(profile);
const oneNoteManifest = surfaces.includes('onenote') ? generatedOneNoteManifestPath(profile) : null;
// Per-app XML manifests are for Upload Add-in testing, so only development ships them.
const officeXmlSurfaces = isDev ? ['word', 'excel', 'powerpoint', 'outlook'] : [];
const officeXmlManifests = officeXmlSurfaces.map((surface) => ({
  surface,
  path: generatedOfficeXmlManifestPath(profile, surface),
}));
const centralizedOfficeManifest = hasDocumentHost
  ? generatedOfficeXmlManifestPath(profile, 'office')
  : null;
const centralizedOutlookManifest = surfaces.includes('outlook')
  ? generatedOfficeXmlManifestPath(profile, 'outlook')
  : null;
const web = join(repoRoot, 'packages', 'web-shell', 'dist-web');
// The brand's icons (brands/<GE_BRAND>/icons), the same files the web build serves.
const iconDir = BRAND.iconsDir;
const requiredIcons = BRAND_ICONS;

const expected = [
  ['manifest', manifest],
  ['OneNote manifest', oneNoteManifest],
  ...officeXmlManifests.map(({ surface, path }) => [`${surface} XML manifest`, path]),
  ['centralized Office XML manifest', centralizedOfficeManifest],
  ['centralized Outlook XML manifest', centralizedOutlookManifest],
];
for (const [what, path] of expected) {
  if (path && !existsSync(path)) {
    console.error(`Generated ${label} ${what} missing: ${path}`);
    process.exit(1);
  }
}
if (!existsSync(web)) {
  console.error(`Built web shell missing: ${web}. Run bun run build first.`);
  process.exit(1);
}
for (const icon of requiredIcons) {
  const iconPath = join(iconDir, icon);
  if (!existsSync(iconPath)) {
    console.error(`${label} icon missing: ${iconPath}`);
    process.exit(1);
  }
}

const outDir = packageDir(profile);
const m365Dir = join(outDir, 'm365');
const oneNoteDir = join(outDir, 'onenote');
const xmlDir = join(outDir, 'xml');
const centralizedDir = join(outDir, 'centralized');
const webDir = join(outDir, 'web');
cleanDir(outDir);

copyFile(manifest, join(m365Dir, 'manifest.json'));
for (const icon of requiredIcons) copyFile(join(iconDir, icon), join(m365Dir, icon));
if (oneNoteManifest) copyFile(oneNoteManifest, join(oneNoteDir, 'onenote.manifest.xml'));
for (const { surface, path } of officeXmlManifests) {
  copyFile(path, join(xmlDir, `${surface}.manifest.xml`));
}
if (centralizedOfficeManifest) {
  copyFile(centralizedOfficeManifest, join(centralizedDir, 'office.manifest.xml'));
}
if (centralizedOutlookManifest) {
  copyFile(centralizedOutlookManifest, join(centralizedDir, 'outlook.manifest.xml'));
}
copyDir(web, webDir);

const commandsChunk = walk(join(webDir, 'assets')).find((file) => /commands-.*\.js$/.test(file));
if (commandsChunk) copyFile(commandsChunk, join(webDir, 'assets', 'commands.js'));

const releaseNotes = isDev
  ? [
      `# ${BRAND.name} Development Sideload Package v${rootVersion()}`,
      '',
      'Profile: development',
      'Unified package surfaces: Word, Excel, PowerPoint, Outlook',
      'Companion package: OneNote legacy XML manifest',
      'Centralized deployment: centralized/office.manifest.xml + centralized/outlook.manifest.xml',
      '',
      'This package is for local/end-to-end development and is not a production alpha artifact.',
      'Run the web shell with `bun run --filter @ge/web-shell dev` while sideloading this package.',
      '',
    ].join('\n')
  : [
      `# ${BRAND.name} Production Package v${rootVersion()}`,
      '',
      `Profile: ${profile}`,
      `Apps: ${surfaces.join(', ')}`,
      'Upload m365/ (unified package) or centralized/ through the Microsoft 365 admin center.',
      oneNoteManifest ? 'OneNote ships separately: onenote/onenote.manifest.xml.' : '',
      '',
    ].join('\n');
writeFileSync(join(outDir, 'README.md'), releaseNotes);

const zipPath = join(repoRoot, 'dist', 'release', `${profile}-m365-v${rootVersion()}.zip`);
createZip(walk(m365Dir), m365Dir, zipPath);
const centralizedZipPath = join(
  repoRoot,
  'dist',
  'release',
  `${profile}-office-centralized-v${rootVersion()}.zip`,
);
createZip(walk(centralizedDir), centralizedDir, centralizedZipPath);

const checksumPath = join(repoRoot, 'dist', 'release', 'SHA256SUMS');
writeChecksums(
  [
    zipPath,
    centralizedZipPath,
    manifest,
    oneNoteManifest,
    centralizedOfficeManifest,
    centralizedOutlookManifest,
    ...officeXmlManifests.map((x) => x.path),
  ].filter(Boolean),
  checksumPath,
);

const m365Manifest = JSON.parse(readFileSync(manifest, 'utf8'));
const officeXml = Object.fromEntries(
  officeXmlManifests.map(({ surface, path }) => [
    surface,
    {
      manifest: path,
      manifestSha256: sha256File(path),
      uploadPath: join(xmlDir, `${surface}.manifest.xml`),
    },
  ]),
);
const sha = (path) => (path ? sha256File(path) : null);
writeJson(join(repoRoot, 'dist', 'release', `${profile}-artifact.json`), {
  profile,
  version: rootVersion(),
  m365Package: zipPath,
  m365PackageSha256: sha256File(zipPath),
  m365Manifest: manifest,
  m365ManifestSha256: sha256File(manifest),
  oneNoteManifest,
  oneNoteManifestSha256: sha(oneNoteManifest),
  officeXml,
  centralizedDeployment: {
    package: centralizedZipPath,
    packageSha256: sha256File(centralizedZipPath),
    officeManifest: join(centralizedDir, 'office.manifest.xml'),
    officeManifestSha256: sha(centralizedOfficeManifest),
    outlookManifest: join(centralizedDir, 'outlook.manifest.xml'),
    outlookManifestSha256: sha(centralizedOutlookManifest),
  },
  manifestVersion: m365Manifest.version,
  webBuild: webDir,
});

console.log(`packaged ${zipPath}`);
console.log(`sha256 ${sha256File(zipPath)}`);
if (oneNoteManifest) console.log(`onenote ${join(oneNoteDir, 'onenote.manifest.xml')}`);
if (officeXmlManifests.length) console.log(`office xml ${xmlDir}`);
console.log(`centralized office xml ${centralizedDir}`);
console.log(`centralized package ${centralizedZipPath}`);
