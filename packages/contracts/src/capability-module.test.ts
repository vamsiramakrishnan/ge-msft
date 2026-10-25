import { describe, expect, it } from 'vitest';
import { capabilityModule, listCapabilityModules } from './capability-module.js';

describe('strict native capability modules', () => {
  it('rejects unrelated effects and unknown nested fields before host access', () => {
    const word = capabilityModule('word', 'tracked-change')!;
    expect(
      word.params.safeParse({ text: 'replacement', mail: { to: ['someone@example.test'] } })
        .success,
    ).toBe(false);
    expect(
      word.params.safeParse({ text: 'replacement', target: { matchText: 'old', range: 'A1' } })
        .success,
    ).toBe(false);
    expect(
      word.params.safeParse({ text: 'replacement', target: { matchText: 'old', injected: true } })
        .success,
    ).toBe(false);
    expect(word.params.safeParse({ target: { matchText: 'old' } }).success).toBe(false);
    expect(word.params.parse({ text: 'replacement', target: { matchText: 'old' } })).toEqual({
      text: 'replacement',
      target: { matchText: 'old' },
    });
  });
  it('keeps canonical nested refinements and rejects ignored conditional-rule fields', () => {
    const excel = capabilityModule('excel', 'format-conditional')!;
    expect(
      excel.params.safeParse({
        conditional: { range: 'Sheet1!A1:A4', rule: { kind: 'dataBar', injected: 'wrong target' } },
      }).success,
    ).toBe(false);
    expect(
      excel.params.safeParse({ conditional: { range: 'Sheet1!A1:A4', rule: { kind: 'dataBar' } } })
        .success,
    ).toBe(true);
    const cells = capabilityModule('excel', 'write-cells')!;
    expect(
      cells.params.safeParse({ cellValues: [[1], [1, 2]], target: { range: 'Sheet1!A1' } }).success,
    ).toBe(false);
  });
  it('does not install handlers for unpromoted catalog operations or other surfaces', () => {
    expect(capabilityModule('outlook', 'create-mail-rule')).toBeUndefined();
    expect(capabilityModule('onenote', 'insert-text')).toBeUndefined();
    const modules = listCapabilityModules();
    expect(new Set(modules.map((module) => `${module.surface}:${module.kind}`)).size).toBe(
      modules.length,
    );
    expect(modules.every((module) => module.postcondition.length > 20)).toBe(true);
  });
});
