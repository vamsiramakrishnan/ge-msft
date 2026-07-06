#!/usr/bin/env node
import {
  alphaManifest,
  developmentManifest,
  ensureDir,
  generatedManifestPath,
  generatedOneNoteManifestPath,
  generatedOfficeXmlManifestPath,
  oneNoteManifest,
  outlookXmlManifest,
  PROD_ONENOTE_BRAND,
  productionManifest,
  profileFromArgs,
  parseArgs,
  releaseConfig,
  repoRoot,
  taskPaneXmlManifest,
  writeJson,
} from './common.mjs';
import { dirname, join } from 'node:path';
import { writeFileSync } from 'node:fs';

const args = parseArgs();
const profile = profileFromArgs(args);

try {
  const cfg = releaseConfig(profile);
  const manifest =
    profile === 'development'
      ? developmentManifest(cfg)
      : profile === 'production'
        ? productionManifest(cfg)
        : alphaManifest(cfg);
  const out = generatedManifestPath(profile);
  ensureDir(dirname(out));
  writeJson(out, manifest);
  console.log(`generated ${out}`);
  if (profile === 'development') {
    const oneNoteOut = generatedOneNoteManifestPath(profile);
    ensureDir(join(repoRoot, 'dist', 'manifests'));
    writeFileSync(oneNoteOut, oneNoteManifest(cfg));
    console.log(`generated ${oneNoteOut}`);
    for (const surface of ['word', 'excel', 'powerpoint']) {
      const xmlOut = generatedOfficeXmlManifestPath(profile, surface);
      writeFileSync(xmlOut, taskPaneXmlManifest(cfg, surface));
      console.log(`generated ${xmlOut}`);
    }
    const outlookOut = generatedOfficeXmlManifestPath(profile, 'outlook');
    writeFileSync(outlookOut, outlookXmlManifest(cfg));
    console.log(`generated ${outlookOut}`);
  }
  if (profile === 'production') {
    // Production covers all surfaces: the unified manifest above (Word, Excel, PowerPoint,
    // Outlook, Teams) plus the companion OneNote legacy XML package.
    const oneNoteOut = generatedOneNoteManifestPath(profile);
    ensureDir(dirname(oneNoteOut));
    writeFileSync(oneNoteOut, oneNoteManifest(cfg, PROD_ONENOTE_BRAND));
    console.log(`generated ${oneNoteOut}`);
  }
} catch (err) {
  if (err?.code === 'BLOCKED_EXTERNAL') {
    console.error(`BLOCKED_EXTERNAL: ${err.message}`);
    process.exit(2);
  }
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
}
