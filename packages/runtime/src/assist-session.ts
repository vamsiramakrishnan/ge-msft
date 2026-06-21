import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ContextKind,
  ContextRef,
  ProvenancePayload,
  SourceRef,
  SseEvent,
  UnitDescriptor,
} from '@ge/contracts';
import { SessionContext, StreamAssistClient } from '@ge/gemini-client';
import type { DocBridge } from './bridge.js';

/**
 * The surface-agnostic assist loop — the analog of Claude's add-in runtime, but grounded
 * on the research unit via streamAssist. It ties a `DocBridge` to `@ge/gemini-client`:
 *
 *   attach context (from the bridge)  →  ask (stream a grounded answer)  →
 *   collect provenance + citations    →  apply (reversible actuation via the bridge)
 *
 * Built once; every bridge plugs in unchanged.
 */
export interface AssistSessionOptions {
  unit: UnitDescriptor;
  /** Default kinds to auto-attach from the bridge before a turn (e.g. ['selection']). */
  autoAttach?: ContextKind[];
}

export class AssistSession {
  readonly context = new SessionContext();
  private session: string | undefined;
  private lastProvenance: ProvenancePayload | undefined;
  private readonly citations: SourceRef[] = [];

  constructor(
    private readonly bridge: DocBridge,
    private readonly client: StreamAssistClient,
    private readonly options: AssistSessionOptions,
  ) {}

  /** Pull attachable context from the bridge and add it to the live session set. */
  async attachContext(kinds?: ContextKind[]): Promise<ContextRef[]> {
    const want = kinds ?? this.options.autoAttach;
    const refs = await this.bridge.listContext();
    const chosen = want ? refs.filter((r) => want.includes(r.kind)) : refs;
    for (const ref of chosen) {
      for (const resolved of await this.bridge.resolveContext(ref)) {
        this.context.add(resolved);
      }
    }
    return chosen;
  }

  /** Detach an attached context object by ref id. */
  detach(id: string): void {
    this.context.remove(id);
  }

  /**
   * Ask a grounded question. Auto-attaches the configured context kinds (once), streams
   * the answer, and records the session id, citations, and provenance as they arrive.
   */
  async *ask(query: string): AsyncGenerator<SseEvent> {
    if (this.options.autoAttach && this.context.size === 0) {
      await this.attachContext(this.options.autoAttach);
    }
    const req = {
      intent: 'assist' as const,
      query,
      unit: { ...this.options.unit, surfaceContext: this.surfaceContext() },
    };
    for await (const event of this.client.stream(req, {
      session: this.session,
      context: this.context.list(),
    })) {
      if (event.type === 'citation') this.citations.push(event.source);
      if (event.type === 'provenance') {
        this.lastProvenance = event.payload;
        this.session = event.payload.sessionId ?? this.session;
      }
      yield event;
    }
  }

  /**
   * Apply a proposed write through the bridge — reversibly and provenanced. The agent's
   * last-turn provenance (agent id, sources, session) is stamped onto the actuation so the
   * host's durable metadata records who/what/why.
   */
  async apply(
    kind: ActuationRequest['kind'],
    params: ActuationParams,
    changeId: string,
  ): Promise<ActuationResult> {
    const request: ActuationRequest = {
      changeId,
      kind,
      surface: this.bridge.surface,
      params,
      ...(this.lastProvenance ? { provenance: this.lastProvenance } : {}),
    };
    return this.bridge.actuate(request);
  }

  get sessionId(): string | undefined {
    return this.session;
  }

  get sources(): SourceRef[] {
    return [...this.citations];
  }

  private surfaceContext(): UnitDescriptor['surfaceContext'] {
    // The runtime sends content via query.parts (SessionContext); the surfaceContext just
    // carries the kind so the engine knows the host. Bridges may enrich it if useful.
    switch (this.bridge.surface) {
      case 'excel':
        return { kind: 'excel' };
      case 'powerpoint':
        return { kind: 'powerpoint' };
      case 'onenote':
        return { kind: 'onenote' };
      case 'teams':
        return { kind: 'teams' };
      case 'outlook':
        return { kind: 'outlook' };
      case 'word':
      default:
        return { kind: 'word' };
    }
  }
}
