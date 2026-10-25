import { describe, expect, it } from 'vitest';
import { NativeTraceStore, observeNativeBatch } from './native-trace.js';

describe('native batch traces', () => {
  it('counts awaited syncs and restores the context method', async () => {
    const sync = async () => undefined;
    const context = { sync };
    const store = new NativeTraceStore();
    expect(
      await observeNativeBatch(
        'excel',
        context,
        async (ctx) => {
          await ctx.sync();
          await ctx.sync();
          return 7;
        },
        store,
      ),
    ).toBe(7);
    expect(context.sync).toBe(sync);
    expect(store.list()).toMatchObject([{ surface: 'excel', syncs: 2, outcome: 'completed' }]);
  });
  it('records uncertainty without recording host error text', async () => {
    const context = {
      sync: async () => {
        throw new Error('private document content');
      },
    };
    const original = context.sync;
    const store = new NativeTraceStore();
    await expect(observeNativeBatch('word', context, (ctx) => ctx.sync(), store)).rejects.toThrow(
      'private',
    );
    expect(context.sync).toBe(original);
    expect(store.list()[0]).toMatchObject({ syncs: 1, outcome: 'unknown' });
    expect(JSON.stringify(store.list())).not.toContain('private');
  });
});
