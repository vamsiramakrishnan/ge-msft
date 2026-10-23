#!/usr/bin/env node
import { existsSync, readFileSync } from 'node:fs';
import {
  prodProfile,
  generatedManifestPath,
  generatedOneNoteManifestPath,
  generatedOfficeXmlManifestPath,
  officeXmlVersion,
  parseArgs,
  profileFromArgs,
  releaseConfig,
  validateGeneratedManifest,
} from './common.mjs';

const args = parseArgs();
const profile = profileFromArgs(args);
const path = generatedManifestPath(profile);

if (!existsSync(path)) {
  console.error(`Generated manifest not found: ${path}`);
  process.exit(1);
}

let manifest;
try {
  manifest = JSON.parse(readFileSync(path, 'utf8'));
} catch (err) {
  console.error(
    `Invalid generated manifest JSON: ${err instanceof Error ? err.message : String(err)}`,
  );
  process.exit(1);
}

const errors = validateGeneratedManifest(manifest, profile);
if (errors.length > 0) {
  for (const error of errors) console.error(error);
  process.exit(1);
}

console.log(`validated ${path}`);

function validateXmlManifest(path, checks) {
  if (!existsSync(path)) return [`Generated XML manifest not found: ${path}`];
  const xml = readFileSync(path, 'utf8');
  const errors = [];
  // Production must not point Office at a local or plain-http origin.
  const forbidden =
    profile === prodProfile
      ? ['REPLACE_', 'example.com', 'localhost', 'DefaultValue="http://', '<AppDomain>http://']
      : ['REPLACE_', 'example.com'];
  for (const token of forbidden) {
    if (xml.includes(token)) errors.push(`${path} contains forbidden token ${token}`);
  }
  if (/\{\{[^}]+\}\}/.test(xml)) {
    errors.push(`${path} contains unresolved template syntax`);
  }
  if (!xml.includes(`<Version>${officeXmlVersion()}</Version>`)) {
    errors.push(`${path} version does not match Office XML manifest version ${officeXmlVersion()}`);
  }
  for (const check of checks) {
    if (!xml.includes(check.text)) errors.push(`${path} ${check.message}`);
  }
  return errors;
}

if (profile === 'development' || profile === prodProfile) {
  const allXmlChecks = [
    {
      path: generatedOneNoteManifestPath(profile),
      checks: [
        { text: '<Host Name="Notebook"', message: 'does not declare the Notebook host' },
        {
          text: '<Permissions>ReadWriteDocument</Permissions>',
          message: 'does not declare ReadWriteDocument',
        },
      ],
    },
    ...[
      ['word', 'Document'],
      ['excel', 'Workbook'],
      ['powerpoint', 'Presentation'],
    ].map(([surface, host]) => ({
      path: generatedOfficeXmlManifestPath(profile, surface),
      checks: [
        { text: `<Host Name="${host}"`, message: `does not declare the ${host} host` },
        {
          text: '<Permissions>ReadWriteDocument</Permissions>',
          message: 'does not declare ReadWriteDocument',
        },
        {
          text: `taskpane.html?host=${surface}`,
          message: `does not point at the ${surface} taskpane URL`,
        },
      ],
    })),
    {
      path: generatedOfficeXmlManifestPath(profile, 'office'),
      checks: [
        { text: '<Host Name="Document"', message: 'does not declare the Document host' },
        { text: '<Host Name="Workbook"', message: 'does not declare the Workbook host' },
        { text: '<Host Name="Presentation"', message: 'does not declare the Presentation host' },
        {
          text: '<Control xsi:type="Button" id="openGeminiWordBtn">',
          message: 'does not declare the Word ribbon command',
        },
        {
          text: '<Control xsi:type="Button" id="openGeminiExcelBtn">',
          message: 'does not declare the Excel ribbon command',
        },
        {
          text: '<Control xsi:type="Button" id="openGeminiPowerPointBtn">',
          message: 'does not declare the PowerPoint ribbon command',
        },
        {
          text: '<WebApplicationInfo>',
          message: 'does not declare the Entra web application identity',
        },
      ],
    },
    {
      path: generatedOfficeXmlManifestPath(profile, 'outlook'),
      checks: [
        { text: '<Host Name="Mailbox"', message: 'does not declare the Mailbox host' },
        {
          text: '<Permissions>ReadWriteMailbox</Permissions>',
          message: 'does not declare ReadWriteMailbox',
        },
        {
          text: 'taskpane.html?host=outlook',
          message: 'does not point at the outlook taskpane URL',
        },
      ],
    },
  ];
  // Production validates only the XML manifests it generates for its chosen apps.
  const prodSurfaces = profile === prodProfile ? releaseConfig(profile).surfaces : [];
  const wanted = (path) => {
    if (profile !== prodProfile) return true;
    if (path.endsWith('.office.manifest.xml'))
      return prodSurfaces.some((s) => ['word', 'excel', 'powerpoint'].includes(s));
    if (path.endsWith('.outlook.manifest.xml')) return prodSurfaces.includes('outlook');
    if (path.endsWith('.onenote.manifest.xml')) return prodSurfaces.includes('onenote');
    return false;
  };
  const xmlChecks = allXmlChecks.filter(({ path }) => wanted(path));
  const xmlErrors = xmlChecks.flatMap(({ path, checks }) => validateXmlManifest(path, checks));
  if (xmlErrors.length > 0) {
    for (const error of xmlErrors) console.error(error);
    process.exit(1);
  }
  for (const { path } of xmlChecks) console.log(`validated ${path}`);
}
