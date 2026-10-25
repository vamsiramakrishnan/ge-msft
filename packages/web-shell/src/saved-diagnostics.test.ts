import { describe, it, expect } from 'vitest';
import { SavedDiagnostics } from './saved-diagnostics.js';

function memoryStorage(): Pick<Storage, 'getItem' | 'setItem' | 'removeItem'> & {
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

const DAY = 24 * 60 * 60 * 1000;

describe('SavedDiagnostics', () => {
  it('saves { question, diagnostic } and lists newest first', () => {
    let t = Date.parse('2026-10-09T10:00:00Z');
    const store = new SavedDiagnostics(memoryStorage(), 'u1', () => t);
    store.save({ id: 'a', question: 'add 3 slides', diagnostic: { surface: 'powerpoint' } });
    t += 1000;
    store.save({ id: 'b', question: 'chart this', diagnostic: { surface: 'excel' } });
    expect(store.list().map((e) => [e.id, e.question])).toEqual([
      ['b', 'chart this'],
      ['a', 'add 3 slides'],
    ]);
    expect(store.list()[1]?.diagnostic).toEqual({ surface: 'powerpoint' });
  });

  it('updates one entry when the same request settles again (plan, then run)', () => {
    let t = 0;
    const store = new SavedDiagnostics(memoryStorage(), 'u1', () => (t += 1000));
    store.save({ id: 'a', question: 'q', diagnostic: { stage: 'plan' } });
    store.save({ id: 'a', question: 'q', diagnostic: { stage: 'run' } });
    expect(store.list()).toHaveLength(1);
    expect(store.list()[0]?.diagnostic).toEqual({ stage: 'run' });
  });

  it('keeps 20 entries and drops ones older than 7 days', () => {
    let t = Date.parse('2026-10-01T00:00:00Z');
    const store = new SavedDiagnostics(memoryStorage(), 'u1', () => t);
    store.save({ id: 'old', question: 'old', diagnostic: {} });
    t += 8 * DAY;
    for (let i = 0; i < 25; i++) {
      t += 1000;
      store.save({ id: `n${i}`, question: `q${i}`, diagnostic: {} });
    }
    const ids = store.list().map((e) => e.id);
    expect(ids).toHaveLength(20);
    expect(ids).not.toContain('old');
    expect(ids[0]).toBe('n24');
  });

  it('stores nothing for an empty question, and survives blocked or corrupt storage', () => {
    const storage = memoryStorage();
    const store = new SavedDiagnostics(storage, 'u1');
    store.save({ id: 'a', question: '   ', diagnostic: {} });
    expect(storage.data.size).toBe(0);
    storage.data.set('ge.saved-diagnostics.v1:u1', '{not json');
    expect(store.list()).toEqual([]);
    const blocked = new SavedDiagnostics(
      {
        getItem: () => {
          throw new Error('blocked');
        },
        setItem: () => {
          throw new Error('blocked');
        },
        removeItem: () => {
          throw new Error('blocked');
        },
      },
      'u1',
    );
    expect(() => blocked.save({ id: 'a', question: 'q', diagnostic: {} })).not.toThrow();
    expect(blocked.list()).toEqual([]);
    expect(() => blocked.clear()).not.toThrow();
  });

  it('clear removes everything', () => {
    const store = new SavedDiagnostics(memoryStorage(), 'u1');
    store.save({ id: 'a', question: 'q', diagnostic: {} });
    store.clear();
    expect(store.list()).toEqual([]);
  });

  it('keeps each signed-in user separate, and stores nothing without a user', () => {
    const storage = memoryStorage();
    new SavedDiagnostics(storage, 'alice').save({ id: 'a', question: 'mine', diagnostic: {} });
    expect(new SavedDiagnostics(storage, 'bob').list()).toEqual([]);
    expect(new SavedDiagnostics(storage, 'alice').list()).toHaveLength(1);
    new SavedDiagnostics(storage, undefined).save({ id: 'x', question: 'q', diagnostic: {} });
    expect([...storage.data.keys()]).toEqual(['ge.saved-diagnostics.v1:alice']);
  });

  it('deletes expired entries from storage, not only from the list', () => {
    const storage = memoryStorage();
    let t = Date.parse('2026-10-01T00:00:00Z');
    const store = new SavedDiagnostics(storage, 'u1', () => t);
    store.save({ id: 'a', question: 'q', diagnostic: {} });
    t += 8 * DAY;
    expect(store.list()).toEqual([]);
    expect(storage.data.size).toBe(0);
  });
});
