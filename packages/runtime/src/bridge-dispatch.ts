import { nativeTraces } from './native-trace.js';
import {
  ActuationKindSchema,
  capabilityModule,
  CAPABILITY_REGISTRY,
  contentHash,
  type CapabilityManifest,
  ActuationRequestSchema,
  ChangeIdSchema,
  type ActuationKind,
  type ActuationRequest,
  type ActuationResult,
  type HostRefusalReason,
  type Surface,
} from '@ge/contracts';

export type BridgeActuationHandler<Host> = (
  host: Host,
  request: ActuationRequest,
) => Promise<ActuationResult>;
export type BridgeActuationHandlers<Host> = Readonly<
  Partial<Record<ActuationKind, BridgeActuationHandler<Host>>>
>;
export interface BridgeDispatch<Host> {
  readonly handledActuations: readonly ActuationKind[];
  dispatch(host: Host, request: ActuationRequest): Promise<ActuationResult>;
}
export interface BridgeDispatchOptions<Host = unknown> {
  /** Reported adapters persist locally; unsupported adapters must not imply durable attribution. */
  provenance: 'reported' | 'unsupported';
  capabilities?: (host: Host) => CapabilityManifest | Promise<CapabilityManifest>;
  /** Preserve a surface's precise read-only preflight refusal without dispatching a write. */
  unavailable?: (host: Host, request: ActuationRequest) => ActuationResult | undefined;
  documentKey?: (host: Host) => string | undefined;
}

const SURFACE_NAMES: Record<Surface, string> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
  outlook: 'Outlook',
  onenote: 'OneNote',
  teams: 'Teams',
};

/**
 * Office.js error codes that reject the WHOLE request before any statement runs, for a reason the
 * user can fix. Batches are not transactional: a code a single statement can raise (AccessDenied,
 * *Conflict, ItemNotFound, ApiNotFound) may follow an earlier statement that landed, so it stays a
 * throw (outcome unknown); bridges refuse those cases themselves before writing (e.g. a protected
 * Excel sheet).
 */
// Empty on purpose: a handler may sync more than once, so even "the whole request was rejected"
// can follow a landed write (security review 2026-10-16). Bridges detect refusals in their own
// read-only pre-checks (Excel: edit mode, protected sheet) where a non-write is certain.
const HOST_REFUSALS: Readonly<Record<string, HostRefusalReason>> = {};

/**
 * Execution authority comes from these host handlers, never descriptive capability metadata.
 * Freeze a private table so discovery and dispatch cannot drift after construction. Validation
 * precedes host access; host errors, inverses, verification and provenance remain host-owned.
 */
export function createBridgeDispatch<Host>(
  surface: Surface,
  handlers: BridgeActuationHandlers<Host>,
  options: BridgeDispatchOptions<Host>,
): BridgeDispatch<Host> {
  const table = Object.freeze({ ...handlers });
  const provenance = options.provenance;
  const handledActuations = Object.freeze(
    Object.keys(table).map((kind) => {
      const parsed = ActuationKindSchema.parse(kind);
      if (typeof table[parsed] !== 'function')
        throw new TypeError(`The ${surface} handler for ${parsed} must be a function.`);
      return parsed;
    }),
  );
  return Object.freeze({
    handledActuations,
    async dispatch(host: Host, input: ActuationRequest): Promise<ActuationResult> {
      const parsed = ActuationRequestSchema.safeParse(input);
      if (!parsed.success) {
        // An invalid envelope cannot name a truthful result. Let the caller handle that schema
        // error rather than inventing a change ID or echoing an invalid actuation kind.
        const changeId = ChangeIdSchema.safeParse(input?.changeId);
        const kind = ActuationKindSchema.safeParse(input?.kind);
        if (!changeId.success || !kind.success) throw parsed.error;
        return {
          ok: false,
          changeId: changeId.data,
          kind: kind.data,
          error: {
            code: 'invalid_request',
            message: 'The actuation request does not match its schema.',
          },
        };
      }
      const request = parsed.data;
      if (request.surface !== surface)
        return {
          ok: false,
          changeId: request.changeId,
          kind: request.kind,
          error: {
            code: 'surface_mismatch',
            message: `The ${surface} bridge cannot execute a ${request.surface} request.`,
          },
        };
      if (!Object.hasOwn(table, request.kind))
        return {
          ok: false,
          changeId: request.changeId,
          kind: request.kind,
          error: {
            code: 'unsupported',
            message: `${SURFACE_NAMES[surface]} bridge cannot ${request.kind}`,
          },
        };
      const module = capabilityModule(surface, request.kind);
      if (module && !module.params.safeParse(input.params).success)
        return {
          ok: false,
          changeId: request.changeId,
          kind: request.kind,
          error: {
            code: 'invalid_request',
            message: 'Parameters do not match this capability contract.',
          },
        };
      if (
        request.params.target?.expectedTextHash &&
        !['set-shape-text', 'fill-content-control'].includes(request.kind)
      )
        return {
          ok: false,
          changeId: request.changeId,
          kind: request.kind,
          error: {
            code: 'unsupported_precondition',
            message: 'This capability cannot verify an expected text hash.',
          },
        };
      if (request.params.target?.expectedDocumentId) {
        const key = options.documentKey?.(host);
        if (
          !key ||
          (await contentHash({ surface, documentKey: key })) !==
            request.params.target.expectedDocumentId
        )
          return {
            ok: false,
            changeId: request.changeId,
            kind: request.kind,
            error: {
              code: 'target_conflict',
              message: 'The approved target belongs to a different or unidentified document.',
            },
          };
      }
      if (options.capabilities) {
        const effective = await options.capabilities(host);
        if (!effective.actuations.some((operation) => operation.kind === request.kind))
          return (
            options.unavailable?.(host, request) ?? {
              ok: false,
              degraded: true,
              changeId: request.changeId,
              kind: request.kind,
              error: {
                code: 'unsupported',
                message:
                  'This operation is unavailable in the current host or mode. Requires ' +
                  (CAPABILITY_REGISTRY.find(
                    (entry) => entry.kind === request.kind && entry.surface === surface,
                  )
                    ?.requirementSets.map((set) => set.name + ' ' + set.minVersion)
                    .join(', ') || 'a supported native API') +
                  '.',
              },
            }
          );
      }
      const started = performance.now();
      let outcome: 'completed' | 'refused' | 'unknown' = 'unknown';
      let result: ActuationResult;
      try {
        result = await table[request.kind]!(host, request);
        outcome = result.ok ? 'completed' : result.recoveryPending ? 'unknown' : 'refused';
      } catch (error) {
        const code = (error as { code?: unknown } | null)?.code;
        const reason = typeof code === 'string' ? HOST_REFUSALS[code] : undefined;
        // Anything else may have failed after a write landed: the caller keeps it unknown.
        if (!reason) throw error;
        return {
          ok: false,
          changeId: request.changeId,
          kind: request.kind,
          error: { code: 'host_refused', reason, message: `Office refused the change (${code}).` },
        };
      } finally {
        nativeTraces.record({
          id: crypto.randomUUID(),
          surface,
          kind: request.kind,
          changeId: request.changeId,
          stage: 'dispatch',
          syncs: 0,
          syncMs: 0,
          durationMs: performance.now() - started,
          outcome,
        });
      }
      return provenance === 'unsupported' && result.ok
        ? {
            ...result,
            ...(request.provenance ? { provenanceUnsupported: true } : { provenanceMissing: true }),
          }
        : result;
    },
  });
}
