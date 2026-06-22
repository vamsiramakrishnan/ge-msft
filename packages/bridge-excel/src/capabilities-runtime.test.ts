import { afterEach, describe, expect, it } from 'vitest';
import { isSet, type RequirementsLike } from './capabilities-runtime.js';

describe('isSet (runtime capability gate)', () => {
  afterEach(() => {
    delete (globalThis as { Office?: unknown }).Office;
  });

  it('returns true when the injected requirements report the set as supported', () => {
    const fake: RequirementsLike = { isSetSupported: (n, v) => n === 'ExcelApi' && v === '1.9' };
    expect(isSet('ExcelApi', '1.9', fake)).toBe(true);
  });

  it('returns false when the injected requirements report the set as unsupported', () => {
    const fake: RequirementsLike = { isSetSupported: () => false };
    expect(isSet('ExcelApi', '1.12', fake)).toBe(false);
  });

  it('passes the exact name and version through to isSetSupported', () => {
    const calls: Array<[string, string | undefined]> = [];
    const fake: RequirementsLike = {
      isSetSupported: (n, v) => {
        calls.push([n, v]);
        return true;
      },
    };
    isSet('ExcelApi', '1.9', fake);
    expect(calls).toEqual([['ExcelApi', '1.9']]);
  });

  it('returns false (never throws) when the host implementation throws', () => {
    const fake: RequirementsLike = {
      isSetSupported: () => {
        throw new Error('host blew up');
      },
    };
    expect(isSet('ExcelApi', '1.9', fake)).toBe(false);
  });

  it('returns false when Office is entirely absent (no injection, no global)', () => {
    expect(isSet('ExcelApi', '1.9')).toBe(false);
  });

  it('returns false when requirements is missing the method', () => {
    expect(isSet('ExcelApi', '1.9', {} as unknown as RequirementsLike)).toBe(false);
  });

  it('falls back to the global Office.context.requirements when not injected', () => {
    (globalThis as { Office?: unknown }).Office = {
      context: {
        requirements: { isSetSupported: (n: string, v: string) => n === 'ExcelApi' && v === '1.9' },
      },
    };
    expect(isSet('ExcelApi', '1.9')).toBe(true);
    expect(isSet('ExcelApi', '1.99')).toBe(false);
  });
});
