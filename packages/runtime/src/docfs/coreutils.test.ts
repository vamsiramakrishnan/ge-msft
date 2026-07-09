import { describe, expect, it } from 'vitest';
import { WorkspaceStore } from '../workspace.js';
import { DocFsRouter } from './router.js';
import { workMount } from './work-mount.js';
import { cat, find, grep, head, ls, wc } from './coreutils.js';

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
});
