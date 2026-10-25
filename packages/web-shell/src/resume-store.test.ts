import { describe, expect, it } from 'vitest';
import { ResumeStore } from './resume-store.js';

function memory(): Pick<Storage, 'getItem' | 'setItem' | 'removeItem'> & {
  data: Map<string, string>;
} {
  const data = new Map<string, string>();
  return {
    data,
    getItem: (k) => data.get(k) ?? null,
    setItem: (k, v) => void data.set(k, v),
    removeItem: (k) => void data.delete(k),
  };
}

describe('ResumeStore', () => {
  it('keeps the request and each landed change until the task ends', () => {
    const store = new ResumeStore(memory(), 'u1');
    store.start('Add 3 slides about work life balance', 'doc1');
    store.landed('Slide added: Understanding balance');
    store.landed('Slide added: Boundaries');
    expect(store.read()).toMatchObject({
      request: 'Add 3 slides about work life balance',
      changes: ['Slide added: Understanding balance', 'Slide added: Boundaries'],
    });
    store.clear();
    expect(store.read()).toBeUndefined();
  });

  it('is per user, stores nothing without one, and forgets after 24 hours', () => {
    const storage = memory();
    let t = Date.parse('2026-10-13T08:00:00Z');
    new ResumeStore(storage, 'alice', () => t).start('mine', 'doc1');
    expect(new ResumeStore(storage, 'alice', () => t).read('doc2')).toBeUndefined();
    expect(new ResumeStore(storage, 'alice', () => t).read('doc1')?.request).toBe('mine');
    expect(new ResumeStore(storage, 'bob', () => t).read()).toBeUndefined();
    new ResumeStore(storage, undefined).start('x', 'doc1');
    expect([...storage.data.keys()]).toEqual(['ge.in-progress.v1:alice']);
    t += 25 * 60 * 60 * 1000;
    expect(new ResumeStore(storage, 'alice', () => t).read()).toBeUndefined();
    expect(storage.data.size).toBe(0);
  });
});
