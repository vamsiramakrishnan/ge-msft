// packages/runtime/src/docfs/index.test.ts
import { describe, expect, it } from 'vitest';
import { WorkspaceStore } from '../workspace.js';
import type { DocBridge } from '../bridge.js';
import { createDocFs, ls } from './index.js';

const bridge = {
  surface: 'excel',
  async captureDocState() {
    return { surface: 'excel', outline: [], inventory: [] };
  },
} as unknown as DocBridge;

describe('createDocFs', () => {
  it('mounts /doc (read-only) and /work', async () => {
    const fs = createDocFs({ bridge, workspace: new WorkspaceStore() });
    const roots = await ls(fs, '/');
    expect(roots).toEqual(['doc/', 'work/']);
  });
});
