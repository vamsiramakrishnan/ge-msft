import { describe, expect, it, vi } from 'vitest';
import { asChangeId } from '@ge/contracts';
import type { CommandLoopEvent } from './assist-session.js';
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
  it('fails certification when an applied edit could not restore tracking mode', async () => {
    const report = await certifyNativeFixture(
      {
        async *runCommandProgram(): AsyncGenerator<CommandLoopEvent> {
          yield {
            type: 'write-result',
            turn: 1,
            changeId: 'fixture',
            result: {
              ok: true,
              changeId: asChangeId('fixture'),
              kind: 'tracked-change',
              verification: { status: 'verified' },
              trackingModeRestoreFailed: true,
            },
          };
          yield { type: 'done', turn: 1, answer: '' };
        },
      },
      {
        surface: 'word',
        getHostProfile: () => ({
          surface: 'word',
          platform: 'test-only',
          build: 'test-only',
          mode: 'write',
          requirements: {},
        }),
      },
      nativeFixture('word'),
      'a'.repeat(40),
      {},
    );
    expect(report.status).toBe('failed');
    expect(report.checks.execution).toBe(false);
    expect(report.checks.postcondition).toBe(true);
  });
  it('requires an exact commit and explicit disposable Excel destination', () => {
    expect(() => nativeFixture('excel')).toThrow('explicit');
    expect(() => nativeFixture('excel', 'A1\nset B1 999')).toThrow('explicit');
  });
});
