// packages/runtime/src/docfs/index.ts
import type { DocFs } from '@ge/contracts';
import type { DocBridge } from '../bridge.js';
import type { WorkspaceStore } from '../workspace.js';
import { DocFsRouter } from './router.js';
import { docMount } from './doc-mount.js';
import { workMount } from './work-mount.js';
import { readOnly } from './mount.js';

export * from './mount.js';
export * from './router.js';
export * from './coreutils.js';
export { docMount } from './doc-mount.js';
export { workMount } from './work-mount.js';

/** Assemble the Phase-1 DocFs: `/doc` (read-only, live document) + `/work` (WorkspaceStore). */
export function createDocFs(opts: { bridge: DocBridge; workspace: WorkspaceStore }): DocFs {
  return new DocFsRouter([readOnly(docMount(opts.bridge)), workMount(opts.workspace)]);
}
