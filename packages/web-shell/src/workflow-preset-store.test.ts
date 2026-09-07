import { describe, expect, it } from 'vitest';
import { WorkflowPresetStore } from './workflow-preset-store.js';

function fixture(raw: string | null = null) {
  const values = new Map<string, string>();
  if (raw !== null) values.set('ge.workflow-settings.v1', raw);
  const storage = {
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => {
      values.set(key, value);
    },
  };
  return { store: new WorkflowPresetStore(storage), values };
}
const preset = {
  schemaVersion: 1,
  recipeId: 'duplicate-rows',
  recipeVersion: 1,
  inputs: { sourceRange: 'Orders!A1:B50' },
};

describe('explicit workflow settings storage', () => {
  it('saves canonical defaults and replaces the same recipe without growing a history', () => {
    const { store } = fixture();
    store.save(preset);
    expect(store.list()[0]?.inputs).toEqual({
      sourceRange: 'Orders!A1:B50',
      headers: true,
      keyColumn: 0,
      caseSensitive: true,
    });
    store.save({ ...preset, inputs: { sourceRange: 'Current!A1:B50' } });
    expect(store.list()).toHaveLength(1);
    expect(store.list()[0]?.inputs.sourceRange).toBe('Current!A1:B50');
    store.remove('duplicate-rows');
    expect(store.list()).toEqual([]);
  });

  it.each([
    { ...preset, recipeVersion: 2 },
    { ...preset, recipeId: 'not-installed' },
    { ...preset, approval: true },
    { ...preset, inputs: { sourceRange: 'Orders!A1:B50', resultId: 'a_111111111111111111111111' } },
    { ...preset, inputs: { sourceRange: 'Orders!A1:B50', accessToken: 'not-a-real-token' } },
  ])('rejects unknown versions, capabilities and authority: %j', (untrusted) => {
    const { store, values } = fixture();
    expect(() => store.save(untrusted)).toThrow();
    expect(values.size).toBe(0);
    expect(() => fixture(JSON.stringify([untrusted])).store.list()).toThrow();
  });

  it('bounds untrusted storage before JSON parsing, rejects duplicates, and supports explicit recovery', () => {
    expect(() => fixture(' '.repeat(32_769)).store.list()).toThrow('too large');
    expect(() => fixture(JSON.stringify(Array(13).fill(preset))).store.list()).toThrow(
      'could not be read',
    );
    const { store } = fixture(JSON.stringify([preset, preset]));
    expect(() => store.list()).toThrow('duplicate workflows');
    store.clear();
    expect(store.list()).toEqual([]);
  });

  it('surfaces disabled or full storage instead of claiming settings were saved', () => {
    const store = new WorkflowPresetStore({
      getItem: () => null,
      setItem: () => {
        throw new Error('Quota exceeded');
      },
    });
    expect(() => store.save(preset)).toThrow('Quota exceeded');
  });
});
