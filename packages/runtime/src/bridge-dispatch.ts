import {
  ActuationKindSchema,
  ActuationRequestSchema,
  ChangeIdSchema,
  type ActuationKind,
  type ActuationRequest,
  type ActuationResult,
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
export interface BridgeDispatchOptions {
  /** Reported adapters persist locally; unsupported adapters must not imply durable attribution. */
  provenance: 'reported' | 'unsupported';
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
 * Execution authority comes from these host handlers, never descriptive capability metadata.
 * Freeze a private table so discovery and dispatch cannot drift after construction. Validation
 * precedes host access; host errors, inverses, verification and provenance remain host-owned.
 */
export function createBridgeDispatch<Host>(
  surface: Surface,
  handlers: BridgeActuationHandlers<Host>,
  options: BridgeDispatchOptions,
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
      const result = await table[request.kind]!(host, request);
      return provenance === 'unsupported' && result.ok
        ? {
            ...result,
            ...(request.provenance ? { provenanceDropped: true } : { provenanceMissing: true }),
          }
        : result;
    },
  });
}
