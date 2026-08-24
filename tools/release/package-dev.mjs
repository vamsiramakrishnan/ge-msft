#!/usr/bin/env node
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  cleanDir,
  copyDir,
  copyFile,
  createZip,
  generatedManifestPath,
  generatedOneNoteManifestPath,
  generatedOfficeXmlManifestPath,
  packageDir,
  repoRoot,
  rootVersion,
  sha256File,
  walk,
  writeChecksums,
  writeJson,
} from './common.mjs';

const profile = 'development';
const manifest = generatedManifestPath(profile);
const oneNoteManifest = generatedOneNoteManifestPath(profile);
const officeXmlSurfaces = ['word', 'excel', 'powerpoint', 'outlook'];
const officeXmlManifests = officeXmlSurfaces.map((surface) => ({
  surface,
  path: generatedOfficeXmlManifestPath(profile, surface),
}));
const centralizedOfficeManifest = generatedOfficeXmlManifestPath(profile, 'office');
const centralizedOutlookManifest = generatedOfficeXmlManifestPath(profile, 'outlook');
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
  console.error(`Generated development manifest missing: ${manifest}`);
  process.exit(1);
}
if (!existsSync(oneNoteManifest)) {
  console.error(`Generated development OneNote manifest missing: ${oneNoteManifest}`);
  process.exit(1);
}
for (const { surface, path } of officeXmlManifests) {
  if (!existsSync(path)) {
    console.error(`Generated development ${surface} XML manifest missing: ${path}`);
    process.exit(1);
  }
}
if (!existsSync(centralizedOfficeManifest)) {
  console.error(
    `Generated development centralized Office XML manifest missing: ${centralizedOfficeManifest}`,
  );
  process.exit(1);
}
if (!existsSync(web)) {
  console.error(`Built web shell missing: ${web}. Run bun run build first.`);
  process.exit(1);
}
for (const icon of requiredIcons) {
  const iconPath = join(publicDir, icon);
  if (!existsSync(iconPath)) {
    console.error(`Development icon missing: ${iconPath}`);
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
for (const icon of requiredIcons) copyFile(join(publicDir, icon), join(m365Dir, icon));
copyFile(oneNoteManifest, join(oneNoteDir, 'onenote.manifest.xml'));
for (const { surface, path } of officeXmlManifests) {
  copyFile(path, join(xmlDir, `${surface}.manifest.xml`));
}
copyFile(centralizedOfficeManifest, join(centralizedDir, 'office.manifest.xml'));
copyFile(centralizedOutlookManifest, join(centralizedDir, 'outlook.manifest.xml'));
copyDir(web, webDir);

const commandsChunk = walk(join(webDir, 'assets')).find((file) => /commands-.*\.js$/.test(file));
if (commandsChunk) copyFile(commandsChunk, join(webDir, 'assets', 'commands.js'));

const releaseNotes = [
  `# Gemini Enterprise Development Sideload Package v${rootVersion()}`,
  '',
  'Profile: development',
  'Unified package surfaces: Word, Excel, PowerPoint, Outlook',
  'Companion package: OneNote legacy XML manifest',
  'Centralized deployment: centralized/office.manifest.xml + centralized/outlook.manifest.xml',
  '',
  'This package is for local/end-to-end development and is not a production alpha artifact.',
  'Run the web shell with `bun run --filter @ge/web-shell dev` while sideloading this package.',
  '',
].join('\n');
writeFileSync(join(outDir, 'README.md'), releaseNotes);

const zipPath = join(repoRoot, 'dist', 'release', `development-m365-v${rootVersion()}.zip`);
createZip(walk(m365Dir), m365Dir, zipPath);
const centralizedZipPath = join(
  repoRoot,
  'dist',
  'release',
  `development-office-centralized-v${rootVersion()}.zip`,
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
  ],
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
writeJson(join(repoRoot, 'dist', 'release', 'development-artifact.json'), {
  profile,
  version: rootVersion(),
  m365Package: zipPath,
  m365PackageSha256: sha256File(zipPath),
  m365Manifest: manifest,
  m365ManifestSha256: sha256File(manifest),
  oneNoteManifest,
  oneNoteManifestSha256: sha256File(oneNoteManifest),
  officeXml,
  centralizedDeployment: {
    package: centralizedZipPath,
    packageSha256: sha256File(centralizedZipPath),
    officeManifest: join(centralizedDir, 'office.manifest.xml'),
    officeManifestSha256: sha256File(centralizedOfficeManifest),
    outlookManifest: join(centralizedDir, 'outlook.manifest.xml'),
    outlookManifestSha256: sha256File(centralizedOutlookManifest),
  },
  manifestVersion: m365Manifest.version,
  webBuild: webDir,
});

console.log(`packaged ${zipPath}`);
console.log(`sha256 ${sha256File(zipPath)}`);
console.log(`onenote ${join(oneNoteDir, 'onenote.manifest.xml')}`);
console.log(`office xml ${xmlDir}`);
console.log(`centralized office xml ${centralizedDir}`);
console.log(`centralized package ${centralizedZipPath}`);
