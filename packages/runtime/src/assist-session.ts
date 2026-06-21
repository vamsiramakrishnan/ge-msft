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
import type { HostEvent, TriggerRegistry } from '@ge/triggers';
import type { DocBridge } from './bridge.js';
import { BRIEF_REF_ID, ContextModel, type CommitMode } from './context-model.js';

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
  /** Optional trigger registry: gates every write (pre-actuation) and audits it (post-actuation). */
  triggers?: TriggerRegistry;
  /** Resume a prior session (persisted in host metadata) — the constructed context survives. */
  resumeSessionId?: string;
}

/** What a `prime` turn asks of the engine: absorb the working context, don't act on it. */
const PRIME_INSTRUCTION =
  'Note the attached working context for our conversation. Acknowledge briefly; take no action.';

export class AssistSession {
  readonly context = new SessionContext();
  /** The event-fed constructor of the working-context brief (see context-model.ts). */
  readonly model: ContextModel;
  private session: string | undefined;
  private lastProvenance: ProvenancePayload | undefined;
  private readonly citations: SourceRef[] = [];

  constructor(
    private readonly bridge: DocBridge,
    private readonly client: StreamAssistClient,
    private readonly options: AssistSessionOptions,
  ) {
    this.model = new ContextModel(bridge.surface);
    this.session = options.resumeSessionId;
  }

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

  /** Attach one specific ref (resolve → add). Backs the context tray's attach-by-chip. */
  async attachRef(ref: ContextRef): Promise<void> {
    for (const resolved of await this.bridge.resolveContext(ref)) {
      this.context.add(resolved);
    }
  }

  /**
   * Ask a grounded question. Auto-attaches the configured context kinds (once), streams
   * the answer, and records the session id, citations, and provenance as they arrive.
   */
  async *ask(query: string): AsyncGenerator<SseEvent> {
    if (this.options.autoAttach && this.context.size === 0) {
      await this.attachContext(this.options.autoAttach);
    }
    // Fold any constructed-but-uncommitted brief so it rides this turn (the lazy commit path).
    // Capture the exact version folded: only notes up to here are on the wire, so only these may
    // be marked resident — notes that arrive mid-stream stay pending for the next turn.
    let foldedVersion: number | undefined;
    if (this.model.hasPending) {
      const brief = this.model.pendingBrief();
      if (brief) {
        for (const entry of brief.entries) this.context.add(entry);
        foldedVersion = brief.version;
      }
    }
    const req = {
      intent: 'assist' as const,
      query,
      unit: { ...this.options.unit, surfaceContext: this.surfaceContext() },
    };
    try {
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
      // Reached only on a fully-consumed stream → the folded brief is now in the session
      // history; mark exactly those notes resident. On a mid-stream throw this is skipped, so
      // the notes stay pending and re-fold next turn (at-least-once, never lost).
      if (foldedVersion !== undefined) this.model.markCommitted(foldedVersion);
    } finally {
      // Always unstage the brief part: on success it is resident; on failure it re-folds next turn.
      this.context.remove(BRIEF_REF_ID);
    }
  }

  /**
   * Feed a host event into the working-context model and commit at the checkpoints it signals.
   * This is the "react to events" path that does NOT run the assistant: most events just
   * construct context (folded into the next turn); a few high-value ones prime it now.
   */
  async ingest(event: HostEvent): Promise<void> {
    const hint = this.model.observe(event);
    if (hint.commit) await this.commit(hint.commit);
  }

  /**
   * Commit the constructed brief to the Gemini Enterprise session.
   *  • `fold` — stage it as a query part for the *next* `ask` (no network call).
   *  • `prime` — send it now as a context-only turn so it is resident before the user asks.
   * Both are no-ops when nothing is pending.
   */
  async commit(mode: CommitMode): Promise<void> {
    const brief = this.model.pendingBrief();
    if (!brief) return;
    if (mode === 'fold') {
      for (const entry of brief.entries) this.context.add(entry);
      return; // marked committed when the next ask() completes
    }
    const req = {
      intent: 'assist' as const,
      query: PRIME_INSTRUCTION,
      unit: { ...this.options.unit, surfaceContext: this.surfaceContext() },
    };
    for await (const event of this.client.stream(req, {
      session: this.session,
      context: brief.entries,
    })) {
      if (event.type === 'provenance') this.session = event.payload.sessionId ?? this.session;
    }
    // Mark exactly the notes that were primed (by version) resident — not any that arrived since.
    this.model.markCommitted(brief.version);
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

    // PreToolUse-style gate: a trigger may veto the write before it lands.
    const triggers = this.options.triggers;
    if (triggers) {
      const gate = await triggers.gate({ type: 'pre-actuation', request });
      if (gate.kind === 'block') {
        return {
          ok: false,
          changeId,
          kind,
          error: { code: 'blocked', message: gate.reason },
        };
      }
    }

    const result = await this.bridge.actuate(request);

    // PostToolUse-style audit hook (fire-and-forget).
    if (triggers) void triggers.dispatch({ type: 'post-actuation', request, result });
    return result;
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
