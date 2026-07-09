// packages/runtime/src/docfs/index.test.ts
import { describe, expect, it } from 'vitest';
import { WorkspaceStore } from '../workspace.js';
import type { DocBridge } from '../bridge.js';
import { createDocFs, ls } from './index.js';

const bridge = {
  surface: 'excel',
  async captureDocState() {
    return { surface: 'excel', outline: [{ level: 1, text: 'Summary' }], inventory: [] };
  },
} as unknown as DocBridge;

describe('createDocFs', () => {
  it('mounts /doc (read-only), /work, and /skills', async () => {
    const fs = createDocFs({ bridge, workspace: new WorkspaceStore() });
    const roots = await ls(fs, '/');
    expect(roots).toEqual(['doc/', 'skills/', 'work/']);
  });

  it('mounts /skills empty when no skillFiles are supplied — present, harmless', async () => {
    const fs = createDocFs({ bridge, workspace: new WorkspaceStore() });
    expect(await fs.readdir('/skills')).toEqual([]);
  });

  it('reads through the /skills mount to the supplied skill file map', async () => {
    const fs = createDocFs({
      bridge,
      workspace: new WorkspaceStore(),
      skillFiles: { 'm365-surface-commander/SKILL.md': '# Surface Commander\n' },
    });
    const v = await fs.readFile('/skills/m365-surface-commander/SKILL.md');
    expect(v?.text).toContain('# Surface Commander');
  });

  it('reads through the /doc mount to the fake DocBridge outline', async () => {
    const fs = createDocFs({ bridge, workspace: new WorkspaceStore() });
    const v = await fs.readFile('/doc/outline.md');
    expect(v?.text).toContain('# Summary');
  });

  it('reads through the /work mount to a real WorkspaceStore round trip', async () => {
    const workspace = new WorkspaceStore();
    workspace.save({ name: 'q3.tsv', sourceLabel: 'test', content: 'a\tb\n1\t2\n', kind: 'tsv' });
    const fs = createDocFs({ bridge, workspace });
    const v = await fs.readFile('/work/q3.tsv');
    expect(v?.text).toBe('a\tb\n1\t2\n');
  });
});
