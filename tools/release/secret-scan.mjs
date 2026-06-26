#!/usr/bin/env node
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import {
  alphaProfile,
  artifactPlaceholderFindings,
  generatedManifestPath,
  generatedOneNoteManifestPath,
  packageDir,
  parseArgs,
  profileFromArgs,
  repoRoot,
  scanForbiddenSecrets,
} from './common.mjs';

const args = parseArgs();
const profile = profileFromArgs(args);
const packageRoot = packageDir(profile);
const artifactRoots = [
  generatedManifestPath(profile),
  ...(profile === 'development' ? [generatedOneNoteManifestPath(profile)] : []),
  packageRoot,
];
const roots = [join(repoRoot, 'packages', 'web-shell', 'dist-web'), ...artifactRoots].filter((p) =>
  existsSync(p),
);

const secretFindings = scanForbiddenSecrets(roots);
const placeholderFindings = artifactPlaceholderFindings(
  artifactRoots.filter((p) => existsSync(p)),
  { allowLocalhost: profile !== alphaProfile },
);

if (secretFindings.length || placeholderFindings.length) {
  for (const f of secretFindings) console.error(`secret-like marker: ${f.file} (${f.pattern})`);
  for (const f of placeholderFindings) console.error(`placeholder/dev marker: ${f}`);
  process.exit(1);
}

console.log(`secret scan passed (${roots.length} artifact root(s))`);
