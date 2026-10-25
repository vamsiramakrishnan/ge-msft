import {
  ActuationRequestSchema,
  capabilityModule,
  formulasForRequest,
  hasScalarFormulaFootprint,
  isUnsafeFormula,
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
  if (
    request.kind === 'write-cells' &&
    formulasForRequest(request).some((row) =>
      row.some((formula) => formula && isUnsafeFormula(formula)),
    )
  )
    return {
      request,
      rejection: {
        ok: false,
        changeId: request.changeId,
        kind: request.kind,
        error: {
          code: 'unsafe_formula',
          message: 'This formula contains an active-content or external-reference operation.',
        },
      },
    };
  if (
    request.kind === 'write-cells' &&
    formulasForRequest(request).some((row) =>
      row.some((formula) => formula && !hasScalarFormulaFootprint(formula)),
    )
  )
    return {
      request,
      rejection: {
        ok: false,
        changeId: request.changeId,
        kind: request.kind,
        error: {
          code: 'unknown_write_footprint',
          message:
            'This formula may spill beyond the reviewed cells. Stage it in a disposable workbook or use bounded deterministic analysis before writing.',
        },
      },
    };
  const module = capabilityModule(surface, request.kind);
  if (module && !module.params.safeParse((raw as ActuationRequest).params).success)
    return {
      request,
      rejection: {
        ok: false,
        changeId: request.changeId,
        kind: request.kind,
        error: {
          code: 'invalid_request',
          message: 'Parameters do not match this capability contract.',
        },
      },
    };
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
