import { describe, expect, it } from 'vitest';
import { asChangeId } from '@ge/contracts';
import { ExecutionLedger } from './execution-ledger.js';

describe('ExecutionLedger', () => {
  it('exports an effect error code and reason, never its message', () => {
    const ledger = new ExecutionLedger();
    ledger.record({
      taskId: 't1',
      surface: 'powerpoint',
      mode: 'command',
      status: 'completed',
      startedAt: '2026-10-02T00:00:00.000Z',
      modelTurns: 1,
      toolCalls: 1,
      effects: [
        {
          ok: false,
          changeId: asChangeId('c1'),
          kind: 'set-shape-text',
          error: {
            code: 'target_conflict',
            reason: 'shape_not_found',
            message: 'Shape "Quarterly secret" was not found',
          },
        },
      ],
    } as Parameters<ExecutionLedger['record']>[0]);
    const [run] = ledger.list();
    expect(run?.effects[0]).toEqual({
      changeId: 'c1',
      kind: 'set-shape-text',
      ok: false,
      errorCode: 'target_conflict',
      errorReason: 'shape_not_found',
    });
    expect(JSON.stringify(run)).not.toContain('Quarterly secret');
  });
});
