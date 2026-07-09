import { describe, expect, it } from 'vitest';
import { WorkspaceStore } from '../workspace.js';
import { DocFsRouter } from './router.js';
import { workMount } from './work-mount.js';
import { cat, find, grep, head, ls, tail, wc } from './coreutils.js';

function fs() {
  const s = new WorkspaceStore();
  s.save({ name: 'a.txt', sourceLabel: 't', content: 'one\ntwo\nthree\n' });
  return new DocFsRouter([workMount(s)]);
}

function fsMulti() {
  const s = new WorkspaceStore();
  s.save({ name: 'a.txt', sourceLabel: 't', content: 'one\ntwo\nthree\n' });
  s.save({ name: 'b.md', sourceLabel: 't', content: '# heading\n' });
  return new DocFsRouter([workMount(s)]);
}

function fsWithContent(name: string, content: string) {
  const s = new WorkspaceStore();
  s.save({ name, sourceLabel: 't', content });
  return new DocFsRouter([workMount(s)]);
}

describe('coreutils', () => {
  it('ls lists a directory', async () => expect(await ls(fs(), '/work')).toEqual(['a.txt']));
  it('cat returns file text', async () =>
    expect((await cat(fs(), '/work/a.txt')).text).toContain('two'));
  it('head returns first n lines', async () =>
    expect((await head(fs(), '/work/a.txt', 2)).lines).toEqual(['one', 'two']));
  it('grep finds matches with line numbers', async () => {
    const r = await grep(fs(), '/work', 'two');
    expect(r.matches[0]).toMatchObject({ path: '/work/a.txt', line: 2, text: 'two' });
  });
  it('wc counts lines/bytes', async () => expect((await wc(fs(), '/work/a.txt')).lines).toBe(3));

  it('find without a glob recursively lists all files under a path', async () => {
    const result = await find(fsMulti(), '/');
    expect(result.sort()).toEqual(['/work/a.txt', '/work/b.md']);
  });

  it('find with a glob filters by leaf name pattern', async () => {
    const result = await find(fsMulti(), '/work', '*.txt');
    expect(result).toEqual(['/work/a.txt']);
  });

  it('find with a glob containing regex metacharacters does not throw and matches literally', async () => {
    const s = new WorkspaceStore();
    s.save({ name: 'data[1].csv', sourceLabel: 't', content: 'x' });
    s.save({ name: 'data[2].csv', sourceLabel: 't', content: 'x' });
    const withBrackets = new DocFsRouter([workMount(s)]);
    await expect(find(withBrackets, '/work', 'data[1].csv')).resolves.toEqual([
      '/work/data[1].csv',
    ]);
  });

  it('cat throws on a missing file', async () => {
    await expect(cat(fs(), '/work/missing.txt')).rejects.toThrow();
  });

  it('head throws on a missing file', async () => {
    await expect(head(fs(), '/work/missing.txt')).rejects.toThrow();
  });

  it('wc throws on a missing file', async () => {
    await expect(wc(fs(), '/work/missing.txt')).rejects.toThrow();
  });

  // tail — the file-level DocFs coreutil (last N lines of a saved artifact or document entry).
  // Distinct from compose.ts's pipeline `tail` transform, which operates on already-materialized
  // Value rows within a `(... | tail 5)` composition — this one reads a DocFs path directly.
  it('tail returns the last n lines of a file (default 10)', async () => {
    const content = Array.from({ length: 15 }, (_, i) => `line ${i + 1}`).join('\n');
    const { lines } = await tail(fsWithContent('notes.md', content), '/work/notes.md');
    expect(lines).toEqual(Array.from({ length: 10 }, (_, i) => `line ${i + 6}`));
  });

  it('tail respects an explicit n', async () => {
    const fsInst = fsWithContent('notes.md', 'a\nb\nc\nd');
    expect((await tail(fsInst, '/work/notes.md', 2)).lines).toEqual(['c', 'd']);
  });

  it('tail on n<=0 returns no lines (not the whole file — slice(-0) === slice(0) is a trap)', async () => {
    const fsInst = fsWithContent('notes.md', 'a\nb\nc\nd');
    expect((await tail(fsInst, '/work/notes.md', 0)).lines).toEqual([]);
  });

  it('tail throws on a missing file (matches head/cat/wc — no such file, not an empty result)', async () => {
    await expect(tail(fs(), '/work/missing.md')).rejects.toThrow();
  });

  it('head and tail agree on line-splitting semantics for the same file', async () => {
    const fsInst = fsWithContent('notes.md', 'one\ntwo\nthree\n');
    // A trailing newline splits into a trailing empty-string "line" — head and tail both see it,
    // neither one special-cases the trailing newline.
    expect((await head(fsInst, '/work/notes.md', 4)).lines).toEqual(['one', 'two', 'three', '']);
    expect((await tail(fsInst, '/work/notes.md', 4)).lines).toEqual(['one', 'two', 'three', '']);
  });
});
