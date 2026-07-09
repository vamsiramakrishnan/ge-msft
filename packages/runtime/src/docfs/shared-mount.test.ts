// packages/runtime/src/docfs/shared-mount.test.ts
import { describe, expect, it } from 'vitest';
import { sharedMount, type SharedStore } from './shared-mount.js';

/** An in-memory fake of the SharedStore port, for exercising the mount without Graph. */
function fakeStore(seed: Record<string, string> = {}): SharedStore {
  const files = new Map(Object.entries(seed));
  return {
    list: () =>
      Promise.resolve([...files.entries()].map(([name, text]) => ({ name, size: text.length }))),
    read: (path) => Promise.resolve(files.get(path)),
    write: (path, content) => {
      files.set(path, content);
      return Promise.resolve();
    },
    remove: (path) => {
      files.delete(path);
      return Promise.resolve();
    },
  };
}

describe('sharedMount', () => {
  it('lists flat files at the root', async () => {
    const fs = sharedMount(fakeStore({ 'notes.txt': 'hello', 'summary.md': '# Sum' }));
    const entries = await fs.readdir('');
    expect(entries).toContainEqual({ name: 'notes.txt', kind: 'file', size: 5 });
    expect(entries).toContainEqual({ name: 'summary.md', kind: 'file', size: 5 });
  });

  it('groups a "/"-containing name into a pseudo-directory, naming convention only', async () => {
    const fs = sharedMount(fakeStore({ 'word/notes.txt': 'hi', 'excel/data.tsv': 'a\tb' }));
    const root = await fs.readdir('');
    expect(root).toContainEqual({ name: 'word', kind: 'dir' });
    expect(root).toContainEqual({ name: 'excel', kind: 'dir' });
    const wordEntries = await fs.readdir('word');
    expect(wordEntries).toContainEqual({ name: 'notes.txt', kind: 'file', size: 2 });
  });

  it('reads a file by its full relative path', async () => {
    const fs = sharedMount(fakeStore({ 'notes.txt': 'hello from excel' }));
    const v = await fs.readFile('notes.txt');
    expect(v?.text).toBe('hello from excel');
  });

  it('returns null reading a missing file', async () => {
    const fs = sharedMount(fakeStore());
    expect(await fs.readFile('nope.txt')).toBeNull();
  });

  it('stat reports a file for a leaf name and a dir for a name prefix', async () => {
    const fs = sharedMount(fakeStore({ 'word/notes.txt': 'hi' }));
    expect(await fs.stat('word/notes.txt')).toMatchObject({ kind: 'file' });
    expect(await fs.stat('word')).toMatchObject({ kind: 'dir' });
    expect(await fs.stat('nope')).toBeNull();
  });

  it('searches file content across all shared files and reports matching lines', async () => {
    const fs = sharedMount(fakeStore({ 'a.txt': 'line one\nfind me here\n', 'b.txt': 'nothing' }));
    const hits = await fs.search('', 'find me');
    expect(hits).toEqual([{ path: 'a.txt', line: 2, text: 'find me here' }]);
  });

  it('search is scoped to names under the given path', async () => {
    const fs = sharedMount(fakeStore({ 'word/a.txt': 'target', 'excel/b.txt': 'target' }));
    const hits = await fs.search('word', 'target');
    expect(hits).toHaveLength(1);
    expect(hits[0]?.path).toBe('word/a.txt');
  });

  it('returns [] searching for an empty pattern', async () => {
    const fs = sharedMount(fakeStore({ 'a.txt': 'hi' }));
    expect(await fs.search('', '')).toEqual([]);
  });

  it('degrades to an empty, harmless mount for a store with nothing shared yet', async () => {
    const fs = sharedMount(fakeStore());
    expect(await fs.readdir('')).toEqual([]);
    expect(await fs.readFile('anything')).toBeNull();
    expect(await fs.stat('anything')).toBeNull();
  });

  it('truncates readFile output to maxBytes', async () => {
    const fs = sharedMount(fakeStore({ 'big.txt': 'x'.repeat(200) }));
    const v = await fs.readFile('big.txt', { maxBytes: 10 });
    expect(v?.truncated).toBe(true);
    expect(v?.bytes).toBeLessThanOrEqual(10);
  });

  it('reflects a write made through the store on the next read (mount is a live view, not a snapshot)', async () => {
    const store = fakeStore();
    const fs = sharedMount(store);
    expect(await fs.readFile('new.txt')).toBeNull();
    await store.write('new.txt', 'just shared');
    expect((await fs.readFile('new.txt'))?.text).toBe('just shared');
  });
});
