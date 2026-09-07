import {
  ActuationRequestSchema,
  type ActuationRequest,
  type ActuationResult,
  type CapabilityManifest,
  type Surface,
} from '@ge/contracts';

export type ActuationAdmission =
  | { request: ActuationRequest; rejection?: never }
  | { request: ActuationRequest; rejection: ActuationResult };

/** Every execution route admits against the same current, effective host capabilities. */
export function admitActuationRequest(
  raw: unknown,
  surface: Surface,
  capabilities: CapabilityManifest,
): ActuationAdmission {
  const request = ActuationRequestSchema.parse(raw);
  const wrongSurface = request.surface !== surface || capabilities.surface !== surface;
  if (
    wrongSurface ||
    !capabilities.actuations.some(
      (entry) => entry.kind === request.kind && entry.surface === surface,
    )
  )
    return {
      request,
      rejection: {
        ok: false,
        changeId: request.changeId,
        kind: request.kind,
        error: {
          code: wrongSurface ? 'surface_mismatch' : 'capability_unavailable',
          message: wrongSurface
            ? 'This write belongs to a different document surface.'
            : 'This write is disabled for the current surface or release profile.',
        },
      },
    };
  return { request };
}
