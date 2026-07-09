import { describe, expect, it } from 'vitest';

import { WorkspaceStore } from './workspace.js';

describe('WorkspaceStore', () => {
  it('stores bounded artifacts and clamps cat previews', () => {
    const store = new WorkspaceStore();
    store.save({
      name: 'long.txt',
      sourceLabel: 'literal',
      content: Array.from({ length: 300 }, (_, i) => `line ${i + 1}`).join('\n'),
    });

    const result = store.cat('long.txt', 500);

    expect(result.workspace).toBe('cat');
    if (result.workspace !== 'cat') throw new Error('expected cat result');
    expect(result.head).toBe(200);
    expect(result.preview.split('\n')).toHaveLength(200);
  });

  it('greps artifacts without re-reading the host', () => {
    const store = new WorkspaceStore();
    store.save({
      name: 'schedule.tsv',
      sourceLabel: "read 'Daily schedule'!B3:I53",
      content: 'Time\tMonday\n08:00\tDeep Work\n08:30\tManager Sync\n',
    });

    const result = store.grep('schedule.tsv', 'manager', 0);

    expect(result.workspace).toBe('grep');
    if (result.workspace !== 'grep') throw new Error('expected grep result');
    expect(result.matches).toEqual([{ line: 3, text: '08:30\tManager Sync' }]);
  });
});

describe('WorkspaceStore — cp/mv/rm artifact lifecycle', () => {
  it('cp duplicates an artifact under a new name with a new id', () => {
    const store = new WorkspaceStore();
    store.save({ name: 'a.tsv', sourceLabel: 'test', content: 'x\ty\n1\t2\n', kind: 'tsv' });
    const result = store.cp('a.tsv', 'b.tsv');
    expect(result.workspace).toBe('cp');
    const b = store.get('b.tsv');
    expect(b?.text).toBe(store.get('a.tsv')!.text);
    expect(b?.id).not.toBe(store.get('a.tsv')!.id);
  });

  it('cp on a missing source returns a corrective error', () => {
    const store = new WorkspaceStore();
    expect(store.cp('nope', 'x')).toHaveProperty('error');
  });

  it('mv renames an artifact in place (same id, old name no longer resolves)', () => {
    const store = new WorkspaceStore();
    store.save({ name: 'a.tsv', sourceLabel: 'test', content: '1', kind: 'tsv' });
    const id = store.get('a.tsv')!.id;
    store.mv('a.tsv', 'b.tsv');
    expect(store.get('b.tsv')?.id).toBe(id);
    expect(store.get('a.tsv')).toBeUndefined();
  });

  it('mv on a missing source returns a corrective error', () => {
    const store = new WorkspaceStore();
    expect(store.mv('nope', 'x')).toHaveProperty('error');
  });

  it('rm deletes an artifact so it no longer resolves by name or id', () => {
    const store = new WorkspaceStore();
    store.save({ name: 'a.tsv', sourceLabel: 'test', content: '1', kind: 'tsv' });
    const id = store.get('a.tsv')!.id;
    store.rm('a.tsv');
    expect(store.get('a.tsv')).toBeUndefined();
    expect(store.get(id)).toBeUndefined();
  });

  it('rm on a missing artifact returns a corrective error', () => {
    const store = new WorkspaceStore();
    expect(store.rm('nope')).toHaveProperty('error');
  });

  it('mv/cp onto an existing destination name overwrites it (last-write-wins, matching save() semantics)', () => {
    const store = new WorkspaceStore();
    store.save({ name: 'a.tsv', sourceLabel: 'test', content: '1', kind: 'tsv' });
    store.save({ name: 'b.tsv', sourceLabel: 'test', content: '2', kind: 'tsv' });
    store.cp('a.tsv', 'b.tsv');
    expect(store.get('b.tsv')?.text).toBe('1');
  });

  it("rm by id also frees the artifact's own name alias", () => {
    const store = new WorkspaceStore();
    store.save({ name: 'a.tsv', sourceLabel: 'test', content: '1', kind: 'tsv' });
    const id = store.get('a.tsv')!.id;
    store.rm(id);
    expect(store.get('a.tsv')).toBeUndefined();
    expect(store.get(id)).toBeUndefined();
  });

  it("mv by id also frees the artifact's own prior name alias", () => {
    const store = new WorkspaceStore();
    store.save({ name: 'a.tsv', sourceLabel: 'test', content: '1', kind: 'tsv' });
    const id = store.get('a.tsv')!.id;
    store.mv(id, 'b.tsv');
    expect(store.get('a.tsv')).toBeUndefined();
    expect(store.get('b.tsv')?.id).toBe(id);
  });

  it('cp counts as a new artifact for evictOldest (32-artifact cap)', () => {
    const store = new WorkspaceStore();
    for (let i = 0; i < 32; i++) {
      store.save({ name: `f${i}.txt`, sourceLabel: 'test', content: String(i) });
    }
    expect(store.list()).toHaveLength(32);
    store.cp('f0.txt', 'f0-copy.txt');
    // The copy pushed the store over the cap; the oldest artifact (f0.txt's original, now
    // unreachable except via its own id since f0.txt's alias moved) was evicted.
    expect(store.list()).toHaveLength(32);
    expect(store.get('f0-copy.txt')).toBeDefined();
  });
});
