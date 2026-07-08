// packages/runtime/src/docfs/work-mount.test.ts
import { describe, expect, it } from 'vitest';
import { WorkspaceStore } from '../workspace.js';
import { workMount } from './work-mount.js';

function seeded() {
  const s = new WorkspaceStore();
  s.save({ name: 'q3.tsv', sourceLabel: 'test', content: 'a\tb\n1\t2\n', kind: 'tsv' });
  return s;
}

describe('workMount', () => {
  it('lists artifacts as files', async () => {
    const m = workMount(seeded());
    const entries = await m.readdir('');
    expect(entries.map((e) => e.name)).toContain('q3.tsv');
    expect(entries[0]!.kind).toBe('file');
  });

  it('reads an artifact by name', async () => {
    const m = workMount(seeded());
    const v = await m.readFile('q3.tsv');
    expect(v?.text).toContain('a\tb');
  });

  it('greps across artifacts and returns the artifact-relative path', async () => {
    const m = workMount(seeded());
    const hits = await m.search('', '1');
    expect(hits[0]).toMatchObject({ path: 'q3.tsv', line: 2 });
  });

  it('returns null for a missing artifact', async () => {
    expect(await workMount(seeded()).readFile('nope')).toBeNull();
  });
});
