// packages/runtime/src/docfs/router.test.ts
import { describe, expect, it } from 'vitest';
import { DocFsRouter } from './router.js';
import type { Mount } from './mount.js';

const work: Mount = {
  prefix: 'work',
  readdir: async (rel) => (rel === '' ? [{ name: 'notes.md', kind: 'file', size: 12 }] : []),
  stat: async (rel) => (rel === 'notes.md' ? { path: '', kind: 'file', size: 12 } : null),
  readFile: async (rel) =>
    rel === 'notes.md' ? { path: '', text: 'hello world\n', bytes: 12, truncated: false } : null,
  // Note: the mount's search returns match paths RELATIVE to the mount (e.g. 'notes.md' for a hit
  // inside notes.md), not '' — the router only prefixes whatever rel path the mount hands back.
  search: async (_rel, p) =>
    p === 'hello' ? [{ path: 'notes.md', line: 1, text: 'hello world' }] : [],
};

describe('DocFsRouter', () => {
  const fs = new DocFsRouter([work]);
  it('lists mount roots at /', async () => {
    expect((await fs.readdir('/')).map((e) => e.name)).toContain('work');
  });
  it('dispatches to the mount and stamps the absolute path', async () => {
    expect((await fs.readdir('/work')).map((e) => e.name)).toEqual(['notes.md']);
    const v = await fs.readFile('/work/notes.md');
    expect(v?.text).toBe('hello world\n');
    expect(v?.path).toBe('/work/notes.md');
    const m = await fs.search('/work', 'hello');
    expect(m).toEqual([{ path: '/work/notes.md', line: 1, text: 'hello world' }]);
  });
  it('throws on an unknown mount', async () => {
    await expect(fs.readdir('/nope')).rejects.toThrow(/unknown mount/i);
  });
  it('returns null reading a missing file', async () => {
    expect(await fs.readFile('/work/missing')).toBeNull();
  });
});
