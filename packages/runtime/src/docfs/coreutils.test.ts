import { describe, expect, it } from 'vitest';
import { WorkspaceStore } from '../workspace.js';
import { DocFsRouter } from './router.js';
import { workMount } from './work-mount.js';
import { cat, grep, head, ls, wc } from './coreutils.js';

function fs() {
  const s = new WorkspaceStore();
  s.save({ name: 'a.txt', sourceLabel: 't', content: 'one\ntwo\nthree\n' });
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
});
