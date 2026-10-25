import { describe, expect, it, vi } from 'vitest';
import { certifyNativeFixture, nativeFixture } from './native-certification.js';

describe('native certification boundary', () => {
  it('reports unavailable without running simulated work as native evidence', async () => {
    const runCommandProgram = vi.fn();
    const report = await certifyNativeFixture(
      { runCommandProgram },
      { surface: 'excel' },
      nativeFixture('excel', 'Sheet1!A1:B2'),
      'a'.repeat(40),
      {},
    );
    expect(report.status).toBe('unavailable');
    expect(report.effects).toEqual([]);
    expect(runCommandProgram).not.toHaveBeenCalled();
  });
  it('requires an exact commit and explicit disposable Excel destination', () => {
    expect(() => nativeFixture('excel')).toThrow('explicit');
    expect(() => nativeFixture('excel', 'A1\nset B1 999')).toThrow('explicit');
  });
});
