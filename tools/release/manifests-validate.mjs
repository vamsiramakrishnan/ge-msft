#!/usr/bin/env node
import { existsSync, readFileSync } from 'node:fs';
import {
  generatedManifestPath,
  generatedOneNoteManifestPath,
  generatedOfficeXmlManifestPath,
  officeXmlVersion,
  parseArgs,
  profileFromArgs,
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
  for (const token of ['REPLACE_', 'example.com']) {
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

if (profile === 'development') {
  const xmlChecks = [
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
  const xmlErrors = xmlChecks.flatMap(({ path, checks }) => validateXmlManifest(path, checks));
  if (xmlErrors.length > 0) {
    for (const error of xmlErrors) console.error(error);
    process.exit(1);
  }
  for (const { path } of xmlChecks) console.log(`validated ${path}`);
}
