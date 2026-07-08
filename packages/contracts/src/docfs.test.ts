// packages/contracts/src/docfs.test.ts
import { describe, expect, it } from 'vitest';
import { parseDocPath } from './docfs.js';

describe('parseDocPath', () => {
  it('splits /<mount>/<rest>', () => {
    expect(parseDocPath('/doc/sheets/Q3.tsv')).toEqual({ mount: 'doc', rel: 'sheets/Q3.tsv' });
    expect(parseDocPath('/work')).toEqual({ mount: 'work', rel: '' });
    expect(parseDocPath('/work/')).toEqual({ mount: 'work', rel: '' });
  });
  it('normalizes and rejects traversal + non-absolute', () => {
    expect(parseDocPath('/doc/a/../b')).toEqual({ mount: 'doc', rel: 'b' });
    expect(() => parseDocPath('doc/x')).toThrow(/absolute/);
    expect(() => parseDocPath('/doc/../../etc')).toThrow(/escape/);
    expect(() => parseDocPath('/')).toThrow(/mount/);
  });
});
