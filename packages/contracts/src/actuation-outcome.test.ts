import { describe, expect, it } from 'vitest';
import { asChangeId } from './brand.js';
import { type ActuationRequest, type ActuationResult } from './capability.js';
import { assessActuationResult, validateActuationResult } from './actuation-outcome.js';

const request: ActuationRequest = {
  surface: 'word',
  kind: 'tracked-change',
  changeId: asChangeId('expected'),
  params: { text: 'Updated text' },
};
const result: ActuationResult = { ok: true, kind: request.kind, changeId: request.changeId };

describe('shared actuation outcome contract', () => {
  it.each([
    [result, 'unverified'],
    [{ ...result, verification: { status: 'verified' } }, 'verified'],
    [{ ...result, verification: { status: 'unknown' } }, 'uncertain'],
    [{ ...result, ok: false, verification: { status: 'mismatch' } }, 'uncertain'],
    [{ ...result, ok: false, recoveryPending: true }, 'uncertain'],
    [
      { ...result, ok: false, error: { code: 'outcome_unknown', message: 'Interrupted' } },
      'uncertain',
    ],
    [{ ...result, ok: false, error: { code: 'unapproved', message: 'Declined' } }, 'rejected'],
    [
      { ...result, ok: false, error: { code: 'precondition_failed', message: 'Changed' } },
      'failed',
    ],
  ] as const)('classifies actual evidence without inventing readback', (receipt, outcome) => {
    expect(assessActuationResult(receipt)).toBe(outcome);
  });

  it.each([
    undefined,
    {},
    { ...result, ok: 'true' },
    { ...result, kind: 'write-cells' },
    {
      ...result,
      verification: { status: 'verified' },
      error: { code: 'stale_source', message: 'Changed' },
    },
    { ...result, ok: false, verification: { status: 'verified' } },
    { ...result, changeId: 'some-other-effect', verification: { status: 'verified' } },
  ])('rejects malformed or uncorrelated post-dispatch receipts as uncertain', (raw) => {
    expect(validateActuationResult(request, raw)).toMatchObject({
      ok: false,
      changeId: request.changeId,
      kind: request.kind,
      recoveryPending: true,
      error: { code: 'outcome_unknown' },
    });
  });

  it('marks an unattributed successful write without inventing provenance', () => {
    expect(validateActuationResult(request, result)).toEqual({
      ...result,
      provenanceMissing: true,
    });
  });
});
