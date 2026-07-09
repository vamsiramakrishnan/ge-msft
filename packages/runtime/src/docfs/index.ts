// packages/runtime/src/docfs/index.ts
import type { DocFs } from '@ge/contracts';
import type { DocBridge } from '../bridge.js';
import type { WorkspaceStore } from '../workspace.js';
import { DocFsRouter } from './router.js';
import { docMount } from './doc-mount.js';
import { workMount } from './work-mount.js';
import { skillsMount } from './skills-mount.js';
import { sharedMount, type SharedStore } from './shared-mount.js';
import { readOnly } from './mount.js';

export * from './mount.js';
export * from './router.js';
export * from './coreutils.js';
export { docMount } from './doc-mount.js';
export { workMount } from './work-mount.js';
export { skillsMount } from './skills-mount.js';
export { sharedMount, type SharedStore } from './shared-mount.js';

/** A `SharedStore` with no backing (Graph consent not granted / not wired) — `/shared` degrades to
 *  an empty, read-only mount rather than a required dependency, mirroring `/skills`'s default. */
const NO_SHARED_STORE: SharedStore = {
  list: () => Promise.resolve([]),
  read: () => Promise.resolve(undefined),
  write: () => Promise.reject(new Error('sharing is not configured for this session')),
  remove: () => Promise.reject(new Error('sharing is not configured for this session')),
};

/**
 * Assemble the DocFs: `/doc` (read-only, live document), `/work` (WorkspaceStore), `/skills`
 * (read-only, the caller-supplied skill-bundle file map), and `/shared` (the cross-surface handoff
 * store — read-only at the DocFs layer; writes happen only via the `share` workspace verb, exactly
 * like `/work`'s `save`/`cp`/`mv`/`rm` sit outside DocFs's own read-only contract).
 */
export function createDocFs(opts: {
  bridge: DocBridge;
  workspace: WorkspaceStore;
  skillFiles?: Readonly<Record<string, string>>;
  sharedStore?: SharedStore;
}): DocFs {
  return new DocFsRouter([
    readOnly(docMount(opts.bridge)),
    workMount(opts.workspace),
    readOnly(skillsMount(opts.skillFiles ?? {})),
    readOnly(sharedMount(opts.sharedStore ?? NO_SHARED_STORE)),
  ]);
}
