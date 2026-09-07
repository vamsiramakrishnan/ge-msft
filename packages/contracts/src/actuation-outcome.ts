import {
  ActuationResultSchema,
  type ActuationRequest,
  type ActuationResult,
} from './capability.js';

/** Application truth is independent of whether a host supports verified readback. */
export type ActuationOutcome = 'verified' | 'unverified' | 'uncertain' | 'rejected' | 'failed';

export function assessActuationResult(result: ActuationResult): ActuationOutcome {
  if (
    result.recoveryPending ||
    (result.ok && result.error) ||
    (!result.ok && result.verification?.status === 'verified') ||
    result.error?.code === 'outcome_unknown' ||
    (result.verification && result.verification.status !== 'verified')
  )
    return 'uncertain';
  if (result.ok) return result.verification?.status === 'verified' ? 'verified' : 'unverified';
  return ['unapproved', 'plan_unapproved'].includes(result.error?.code ?? '')
    ? 'rejected'
    : 'failed';
}

/** Use only after dispatch: the host may have changed the document before losing its receipt. */
export function unknownActuationResult(
  request: ActuationRequest,
  message: string,
): ActuationResult {
  return {
    ok: false,
    changeId: request.changeId,
    kind: request.kind,
    recoveryPending: true,
    error: { code: 'outcome_unknown', message },
  };
}

/** Parse and correlate a host receipt before it can affect recovery, completion or attribution. */
export function validateActuationResult(request: ActuationRequest, raw: unknown): ActuationResult {
  const parsed = ActuationResultSchema.safeParse(raw);
  if (
    !parsed.success ||
    parsed.data.kind !== request.kind ||
    parsed.data.changeId !== request.changeId
  )
    return unknownActuationResult(
      request,
      'The host returned an invalid or mismatched write receipt. Inspect the document before retrying.',
    );
  const result = parsed.data;
  if ((result.ok && result.error) || (!result.ok && result.verification?.status === 'verified'))
    return unknownActuationResult(
      request,
      'The host returned contradictory write evidence. Inspect the document before retrying.',
    );
  // An absent attribution payload must never look like a durably attributed write.
  return result.ok && !request.provenance ? { ...result, provenanceMissing: true } : result;
}
