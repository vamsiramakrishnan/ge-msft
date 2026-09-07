import { describe, expect, it } from 'vitest';
import { CommandResultStore } from './result-store.js';

const bytes = (text: string): number => new TextEncoder().encode(text).byteLength;
const encoded = (store: CommandResultStore, values: unknown[]): Array<Record<string, unknown>> =>
  JSON.parse(store.encode(values).text) as Array<Record<string, unknown>>;
const receiptRef = (store: CommandResultStore, value: unknown): string => {
  const receipt = encoded(store, [value])[0]!;
  expect(receipt.ref).toEqual(expect.any(String));
  return receipt.ref as string;
};
const inspect = (store: CommandResultStore, selector: string): Record<string, unknown> =>
  store.inspect(selector) as Record<string, unknown>;

describe('CommandResultStore', () => {
  it('keeps small results inline without changing result shapes', () => {
    const store = new CommandResultStore();
    const values = [{ ok: true, answer: 42 }, { error: 'Target moved' }, [1, null, 'x']];
    const result = store.encode(values);
    expect(JSON.parse(result.text)).toEqual(values);
    expect(result).toMatchObject({ retained: 0, errors: 0, inputBytesComplete: true });
    expect(result.inputBytes).toBe(bytes(JSON.stringify(values)));
    expect(result.outputBytes).toBe(bytes(result.text));
  });

  it('stores detached payloads and preserves safety-critical result metadata', () => {
    const store = new CommandResultStore({ inlineBytes: 100 });
    const value = {
      ok: true,
      verification: { status: 'mismatch', reason: 'Coauthor edit' },
      recoveryPending: true,
      truncated: true,
      error: 'Readback failed',
      data: Array.from({ length: 50 }, (_, i) => i),
    };
    const receipt = encoded(store, [value])[0]!;
    expect(receipt).toMatchObject({
      ok: true,
      verification: value.verification,
      recoveryPending: true,
      truncated: true,
      error: 'Readback failed',
      complete: false,
    });
    value.data[0] = 900;
    const page = inspect(store, `${receipt.ref} path=/data offset=0 limit=3`);
    expect(page).toMatchObject({ preview: [0, 1, 2], count: 3, total: 50, complete: false });
    (page.preview as number[])[0] = 1000;
    expect(inspect(store, `${receipt.ref} path=/data limit=1`).preview).toEqual([0]);
  });

  it('includes bounded structural hints without disclosing payload values', () => {
    const store = new CommandResultStore({ inlineBytes: 1 });
    const object = Object.fromEntries(
      Array.from({ length: 20 }, (_, index) => [`field${index}`, 'private value']),
    );
    const receipt = encoded(store, [object])[0]!;
    expect(receipt).toMatchObject({
      keys: Object.keys(object).slice(0, 12),
      keyCount: 20,
      keysComplete: false,
    });
    expect(JSON.stringify(receipt)).not.toContain('private value');
    expect(encoded(store, [[1, 2, 3]])[0]).toHaveProperty('length', 3);
    expect(encoded(store, ['A🌐B'])[0]).toHaveProperty('length', 3);
    const longKey = encoded(store, [{ ['x'.repeat(2000)]: 1, data: 2 }])[0]!;
    expect(longKey).toMatchObject({ keys: ['data'], keyCount: 2, keysComplete: false });
  });

  it('budgets UTF-8 bytes across the whole response', () => {
    const store = new CommandResultStore({ turnBytes: 1024, inlineBytes: 512 });
    const values = Array.from({ length: 12 }, (_, index) => ({ index, text: '🌐'.repeat(90) }));
    const result = store.encode(values);
    expect(result.outputBytes).toBe(bytes(result.text));
    expect(result.outputBytes).toBeLessThanOrEqual(1024);
    expect(result.inputBytes).toBe(bytes(JSON.stringify(values)));
    expect(result.retained).toBeGreaterThan(0);
    const first = (JSON.parse(result.text) as Array<Record<string, unknown>>)[0]!;
    expect(first.ref).toBeDefined();
  });

  it('paginates exact array elements and escaped JSON Pointer keys', () => {
    const store = new CommandResultStore({ inlineBytes: 1 });
    const ref = receiptRef(store, { 'a/b': { '~key': [0, 1, 2, 3, 4] } });
    const first = inspect(store, `${ref} path=/a~1b/~0key limit=2`);
    expect(first).toMatchObject({ preview: [0, 1], count: 2, total: 5, complete: false });
    const second = inspect(store, first.next as string);
    expect(second.preview).toEqual([2, 3]);
    expect(inspect(store, second.next as string)).toMatchObject({
      preview: [4],
      offset: 4,
      count: 1,
    });
    expect(inspect(store, `${ref} path=/a~1b/~0key limit=5`)).toMatchObject({
      complete: true,
      preview: [0, 1, 2, 3, 4],
    });
    expect(inspect(store, `${ref} path=/a~1b/~0key offset=5`)).toMatchObject({
      preview: [],
      count: 0,
    });
    expect(inspect(store, `${ref} path=/a~1b/~0key offset=6`)).toHaveProperty('error');
  });

  it('uses Unicode code-point string offsets', () => {
    const store = new CommandResultStore({ inlineBytes: 1 });
    const ref = receiptRef(store, 'A🦊B🌐C');
    const first = inspect(store, `${ref} offset=1 limit=2`);
    expect(first).toMatchObject({ preview: '🦊B', total: 5, count: 2 });
    expect(inspect(store, first.next as string).preview).toBe('🌐C');
  });

  it('inspects small windows of a multi-megabyte string in either direction', () => {
    const store = new CommandResultStore();
    const prefix = 'a'.repeat(7 * 1024 * 1024);
    const ref = receiptRef(store, `${prefix}🌐文🦊tail`);
    const tail = inspect(store, `${ref} offset=${prefix.length} limit=2`);
    expect(tail).toMatchObject({ preview: '🌐文', total: prefix.length + 7, count: 2 });
    expect(inspect(store, tail.next as string)).toMatchObject({ preview: '🦊t', count: 2 });
    expect(inspect(store, `${ref} limit=1`)).toMatchObject({ preview: 'a', count: 1 });
    expect(inspect(store, `${ref} offset=${prefix.length + 2} limit=5`)).toMatchObject({
      preview: '🦊tail',
      count: 5,
    });
  });

  it.each([2048, 4096])(
    'keeps inspected array and Unicode pages inline at a %i-byte threshold',
    (inlineBytes) => {
      const store = new CommandResultStore({ inlineBytes });
      const rows = Array.from({ length: 200 }, (_, index) => ({ index, text: '文🌐'.repeat(10) }));
      const ref = receiptRef(store, rows);
      let selector = `${ref} limit=200`;
      const collected: unknown[] = [];
      while (selector) {
        const page = inspect(store, selector);
        const result = store.encode([page]);
        expect(result.retained).toBe(0);
        expect(bytes(JSON.stringify(page))).toBeLessThanOrEqual(inlineBytes);
        const returned = JSON.parse(result.text)[0] as Record<string, unknown>;
        expect(returned).toEqual(page);
        expect(returned.preview).toEqual(expect.any(Array));
        collected.push(...(returned.preview as unknown[]));
        selector = (returned.next as string) ?? '';
      }
      expect(collected).toEqual(rows);

      const text = '文🌐'.repeat(2000);
      const textRef = receiptRef(store, text);
      const page = inspect(store, `${textRef} limit=200`);
      const result = store.encode([page]);
      expect(result.retained).toBe(0);
      expect(JSON.parse(result.text)[0]).toMatchObject({
        ref: textRef,
        preview: '文🌐'.repeat(100),
        count: 200,
      });
      expect(inspect(store, page.next as string)).toMatchObject({
        preview: '文🌐'.repeat(100),
        offset: 200,
      });
    },
  );

  it('bounds explicit inspection budgets by the shared turn cap', () => {
    const store = new CommandResultStore({ turnBytes: 1024, inlineBytes: 1 });
    const ref = receiptRef(
      store,
      Array.from({ length: 200 }, (_, index) => ({ index, text: '🌐'.repeat(30) })),
    );
    expect(bytes(JSON.stringify(store.inspect(`${ref} limit=200`, 16_384)))).toBeLessThanOrEqual(
      1024,
    );
    expect(store.inspect(ref, NaN)).toHaveProperty('code', 'inspection-budget');
  });

  it('supports quoted JSON Pointers and bounds long selector envelopes', () => {
    const store = new CommandResultStore({ turnBytes: 1024, inlineBytes: 1 });
    const key = 'x'.repeat(1500);
    const ref = receiptRef(store, { 'a b': [1, 2, 3], [key]: [] });
    const first = inspect(store, `${ref} path="/a b" limit=1`);
    expect(first.preview).toEqual([1]);
    expect(inspect(store, first.next as string).preview).toEqual([2]);
    const large = inspect(store, `${ref} path=/${key}`);
    expect(bytes(JSON.stringify(large))).toBeLessThanOrEqual(1024);
    expect(large).toHaveProperty('error');
  });

  it('requires deliberate projection for objects and oversized array elements', () => {
    const store = new CommandResultStore({ turnBytes: 1024, inlineBytes: 1 });
    const ref = receiptRef(store, { rows: [{ body: 'x'.repeat(10_000) }], status: 'ready' });
    expect(inspect(store, ref)).toMatchObject({
      projection: 'keys',
      preview: ['rows', 'status'],
      complete: false,
    });
    expect(inspect(store, `${ref} path=/rows`)).toMatchObject({
      complete: false,
      nextPath: '/rows/0',
      count: 0,
    });
    expect(inspect(store, `${ref} path=/rows/0/body limit=20`)).toMatchObject({
      preview: 'x'.repeat(20),
    });
  });

  it('rejects foreign, expired, and fabricated references', () => {
    const store = new CommandResultStore({ inlineBytes: 1 });
    const other = new CommandResultStore({ inlineBytes: 1 });
    const old = receiptRef(store, { text: 'hello' });
    expect(inspect(other, old)).toHaveProperty('error');
    expect(inspect(store, old.replace(/:\d+$/, ':999'))).toHaveProperty('error');
    store.clear();
    const current = receiptRef(store, { text: 'hello' });
    expect(current).not.toBe(old);
    expect(inspect(store, old)).toHaveProperty('error');
    expect(inspect(store, current)).not.toHaveProperty('error');
  });

  it('reports item, total storage, and item-count quota failures explicitly', () => {
    const store = new CommandResultStore({
      inlineBytes: 1,
      itemBytes: 120,
      totalBytes: 180,
      maxItems: 2,
    });
    const first = receiptRef(store, 'a'.repeat(100));
    expect(encoded(store, ['b'.repeat(100)])[0]).toHaveProperty(
      'storageError.code',
      'store-budget',
    );
    expect(encoded(store, ['c'.repeat(200)])[0]).toHaveProperty('storageError.code', 'item-budget');
    expect(inspect(store, first)).not.toHaveProperty('error');
    const countLimited = new CommandResultStore({ inlineBytes: 1, maxItems: 1 });
    receiptRef(countLimited, [1, 2, 3]);
    expect(encoded(countLimited, [[4, 5]])[0]).toHaveProperty('storageError.code', 'store-budget');
  });

  it('fails boundedly on cycles, depth and node excess, without invoking getters', () => {
    const store = new CommandResultStore({ maxDepth: 4, maxNodes: 20 });
    const cyclic: Record<string, unknown> = { ok: true };
    cyclic.self = cyclic;
    expect(encoded(store, [cyclic])[0]).toMatchObject({
      ok: true,
      storageError: { code: 'cyclic-value' },
    });
    expect(encoded(store, [[[[[[[1]]]]]]])[0]).toHaveProperty('storageError.code', 'depth-budget');
    expect(encoded(store, [Array.from({ length: 100 }, (_, i) => i)])[0]).toHaveProperty(
      'storageError.code',
      'node-budget',
    );
    let accessed = false;
    const getter = {
      get data() {
        accessed = true;
        throw new Error('not called');
      },
    };
    expect(encoded(store, [getter])[0]).toHaveProperty('storageError.code', 'accessor-value');
    expect(accessed).toBe(false);
  });

  it('does not call toJSON or include exception stacks', () => {
    const store = new CommandResultStore();
    const error = new Error('Operation failed');
    error.stack = 'private host data'.repeat(100_000);
    expect(encoded(store, [error])).toEqual([{ error: 'Operation failed' }]);
    let called = false;
    expect(
      encoded(store, [
        {
          toJSON() {
            called = true;
            return 'leak';
          },
        },
      ])[0],
    ).toHaveProperty('storageError.code', 'unsupported-value');
    expect(called).toBe(false);
  });

  it('rejects prototype traversal and selector injection', () => {
    const store = new CommandResultStore({ inlineBytes: 1 });
    const ref = receiptRef(
      store,
      JSON.parse('{"constructor":{"prototype":{}},"rows":[1,2],"__proto__":{}}'),
    );
    for (const suffix of [
      'path=/constructor',
      'path=/__proto__',
      'path=/rows/length',
      'path=/rows/01',
      'path=/rows/~2',
      'offset=-1',
      'offset=1.2',
      'limit=0',
      'limit=201',
      'limit=2 limit=3',
      'path=/rows;done',
      'eval=alert(1)',
      'offset=9007199254740992',
    ]) {
      expect(inspect(store, `${ref} ${suffix}`), suffix).toHaveProperty('error');
    }
  });

  it('exposes a retained index when result count exceeds the turn budget and storage is full', () => {
    const store = new CommandResultStore({ turnBytes: 1024, inlineBytes: 1, maxItems: 2 });
    const result = store.encode(
      Array.from({ length: 30 }, (_, i) => ({
        ok: false,
        error: `failure ${i}`,
        body: 'payload'.repeat(20),
      })),
    );
    expect(result.outputBytes).toBeLessThanOrEqual(1024);
    const summary = JSON.parse(result.text)[0] as Record<string, unknown>;
    expect(summary).toHaveProperty('storageError.code', 'turn-budget');
    expect(inspect(store, summary.ref as string)).toMatchObject({ total: 2, count: 2 });
    expect(result.errors).toBeGreaterThan(0);
  });

  it('reports incomplete input accounting and validates configuration', () => {
    const store = new CommandResultStore({ itemBytes: 20 });
    expect(store.encode(['x'.repeat(100)]).inputBytesComplete).toBe(false);
    expect(() => new CommandResultStore({ turnBytes: 10 })).toThrow(RangeError);
    expect(() => new CommandResultStore({ maxNodes: NaN })).toThrow(RangeError);
  });
});
