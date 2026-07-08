// packages/runtime/src/docfs/readonly.test.ts
import { describe, expect, it } from 'vitest';
import { readOnly, type Mount } from './mount.js';

const base: Mount = {
  prefix: 'x',
  readdir: async () => [{ name: 'a', kind: 'file' }],
  stat: async () => null,
  readFile: async () => null,
  search: async () => [],
};

describe('readOnly', () => {
  it('passes reads through and flags readonly', async () => {
    const ro = readOnly(base);
    expect(ro.readonly).toBe(true);
    expect(ro.prefix).toBe('x');
    expect(await ro.readdir('')).toEqual([{ name: 'a', kind: 'file' }]);
  });
});
