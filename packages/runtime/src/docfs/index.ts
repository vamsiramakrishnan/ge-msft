// packages/runtime/src/docfs/index.ts
import type { DocFs } from '@ge/contracts';
import type { DocBridge } from '../bridge.js';
import type { WorkspaceStore } from '../workspace.js';
import { DocFsRouter } from './router.js';
import { docMount } from './doc-mount.js';
import { workMount } from './work-mount.js';
import { skillsMount } from './skills-mount.js';
import { readOnly } from './mount.js';

export * from './mount.js';
export * from './router.js';
export * from './coreutils.js';
export { docMount } from './doc-mount.js';
export { workMount } from './work-mount.js';
export { skillsMount } from './skills-mount.js';

/**
 * Assemble the DocFs: `/doc` (read-only, live document), `/work` (WorkspaceStore), and `/skills`
 * (read-only, the caller-supplied skill-bundle file map — empty when the caller has none to offer,
 * so `/skills` degrades to an empty, harmless mount rather than a required dependency).
 */
export function createDocFs(opts: {
  bridge: DocBridge;
  workspace: WorkspaceStore;
  skillFiles?: Readonly<Record<string, string>>;
}): DocFs {
  return new DocFsRouter([
    readOnly(docMount(opts.bridge)),
    workMount(opts.workspace),
    readOnly(skillsMount(opts.skillFiles ?? {})),
  ]);
}
