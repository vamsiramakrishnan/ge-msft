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

  it('applies the max cap to the combined match list across all artifacts, not per-artifact', async () => {
    const s = new WorkspaceStore();
    s.save({ name: 'a.txt', sourceLabel: 'test', content: 'hit\nhit\nhit\n', kind: 'text' });
    s.save({ name: 'b.txt', sourceLabel: 'test', content: 'hit\nhit\nhit\n', kind: 'text' });
    const hits = await workMount(s).search('', 'hit', { max: 2 });
    expect(hits).toHaveLength(2);
  });

  it('truncates readFile output to maxBytes', async () => {
    const s = new WorkspaceStore();
    const long = 'abcdefghij';
    s.save({ name: 'long.txt', sourceLabel: 'test', content: long, kind: 'text' });
    const v = await workMount(s).readFile('long.txt', { maxBytes: 4 });
    expect(v?.truncated).toBe(true);
    expect(v?.text).toHaveLength(4);
    expect(v?.text).toBe('abcd');
    expect(v?.bytes).toBe(4);
  });

  it('stats an existing artifact and returns null for a missing one', async () => {
    const s = seeded();
    const artifact = s.get('q3.tsv')!;
    const m = workMount(s);
    expect(await m.stat('q3.tsv')).toEqual({ path: '', kind: 'file', size: artifact.bytes });
    expect(await m.stat('nope')).toBeNull();
  });
});
