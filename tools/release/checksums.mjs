#!/usr/bin/env node
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import {
  packageZip,
  parseArgs,
  profileFromArgs,
  repoRoot,
  verifyChecksums,
  walk,
  writeChecksums,
} from './common.mjs';

const args = parseArgs();
const profile = profileFromArgs(args);
const checksumPath = join(repoRoot, 'dist', 'release', 'SHA256SUMS');

if (args.verify) {
  const failures = verifyChecksums(checksumPath);
  if (failures.length) {
    for (const f of failures) console.error(`${f.file}: ${f.error}`);
    process.exit(1);
  }
  console.log(`verified ${checksumPath}`);
  process.exit(0);
}

const files = [packageZip(profile), ...walk(join(repoRoot, 'dist', 'manifests'))].filter((p) =>
  existsSync(p),
);
writeChecksums(files, checksumPath);
console.log(`wrote ${checksumPath}`);
