import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ChangeId,
  ContextKind,
  ContextRef,
  ProvenancePayload,
  ResolvedContext,
  SessionId,
  SourceRef,
  SseEvent,
  UnitDescriptor,
} from '@ge/contracts';
import { estimateTokens } from '@ge/content';
import { SessionContext, StreamAssistClient } from '@ge/gemini-client';
import type { HostEvent, TriggerRegistry } from '@ge/triggers';
import type { DocBridge } from './bridge.js';
import { BRIEF_REF_ID, ContextModel, type CommitMode } from './context-model.js';

/**
 * Bounded-history compaction policy for the resident `SessionContext` (ADR-0003, element 5).
 * Long working sessions accrete attached context (selections, comments, prior briefs) turn
 * over turn; left unbounded, every turn re-ships the whole set as `query.parts[]`. Compaction
 * evicts the oldest low-value entries down to budget while preserving the grounding anchors,
 * the most recent turns, and anything still pending. A no-op until a threshold is exceeded.
 */
export interface CompactionOptions {
  /** Compact once the resident set exceeds this many context parts. */
  maxParts?: number;
  /** Compact once the resident set's estimated tokens exceed this. */
  maxTokens?: number;
  /** Always retain at least this many of the most-recently-attached entries (recent turns). */
  keepRecent?: number;
}

/** Defaults: generous enough to be inert in normal use; trip only on genuinely long sessions. */
const DEFAULT_COMPACTION: Required<CompactionOptions> = {
  maxParts: 64,
  maxTokens: 24_000,
  keepRecent: 8,
};

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
  resumeSessionId?: SessionId;
  /** Bounded-history compaction of the resident context set (ADR-0003 §5). Defaults applied. */
  compaction?: CompactionOptions;
}

/**
 * Grounding anchors — named references the engine resolves against connected data stores or
 * Drive, and person references. These are the cheap, high-value handles the unit is built on;
 * they are never evicted by compaction (preserving them is the whole point of grounding).
 */
const GROUNDING_VALUE_KINDS = new Set(['indexed-document', 'drive-document', 'person']);

/** What a `prime` turn asks of the engine: absorb the working context, don't act on it. */
const PRIME_INSTRUCTION =
  'Note the attached working context for our conversation. Acknowledge briefly; take no action.';

export class AssistSession {
  readonly context = new SessionContext();
  /** The event-fed constructor of the working-context brief (see context-model.ts). */
  readonly model: ContextModel;
  private session: SessionId | undefined;
  private lastProvenance: ProvenancePayload | undefined;
  private readonly citations: SourceRef[] = [];
  private readonly compaction: Required<CompactionOptions>;

  constructor(
    private readonly bridge: DocBridge,
    private readonly client: StreamAssistClient,
    private readonly options: AssistSessionOptions,
  ) {
    this.model = new ContextModel(bridge.surface);
    this.session = options.resumeSessionId;
    this.compaction = { ...DEFAULT_COMPACTION, ...options.compaction };
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
  async *ask(query: string, opts: { signal?: AbortSignal } = {}): AsyncGenerator<SseEvent> {
    if (this.options.autoAttach && this.context.size === 0) {
      await this.attachContext(this.options.autoAttach);
    }
    // Bound the resident set before it goes on the wire. Threshold-guarded inside, so this is a
    // no-op until the session is genuinely large. Runs before the brief fold so a freshly folded
    // pending brief is never a compaction target.
    this.compact();
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
        ...(opts.signal ? { signal: opts.signal } : {}),
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
   * Bound the resident context set (ADR-0003 §5). When the attached set exceeds the configured
   * `maxParts`/`maxTokens` threshold, evict the **oldest, lowest-value** entries down to budget
   * while preserving, unconditionally:
   *   (a) grounding anchors — `indexed-document` / `drive-document` / `person` references (the
   *       unit's cheap, high-value handles), so grounding never degrades;
   *   (b) the most recent `keepRecent` turns — the newest entries by attach order; and
   *   (c) anything still pending — the constructed brief part (never dropped uncommitted).
   *
   * The policy is deliberately simple: oldest-first eviction of evictable (non-preserved) text
   * entries until the set is at or under budget — no similarity/summarisation scheme. Idempotent
   * and threshold-guarded, so it is a no-op until the session is actually large. Returns the
   * number of entries evicted (0 when under threshold), for tests/telemetry.
   */
  compact(): number {
    const entries = this.context.list();
    if (entries.length === 0) return 0;

    const totalTokens = (set: ResolvedContext[]): number =>
      set.reduce((sum, c) => sum + tokensOf(c), 0);

    // Under both thresholds → nothing to do.
    const overParts = entries.length > this.compaction.maxParts;
    const overTokens = totalTokens(entries) > this.compaction.maxTokens;
    if (!overParts && !overTokens) return 0;

    // Partition into preserved (never evicted) and evictable. `list()` is insertion-ordered, so
    // the last `keepRecent` entries are "the most recent turns"; everything in GROUNDING_VALUE_KINDS
    // is an anchor; the pending brief part is preserved by id.
    const recentCutoff = entries.length - this.compaction.keepRecent;
    const isPreserved = (c: ResolvedContext, index: number): boolean =>
      c.ref.id === BRIEF_REF_ID || GROUNDING_VALUE_KINDS.has(c.value.as) || index >= recentCutoff;

    // Oldest evictable first (insertion order). Drop until at/under budget, or until none left.
    const evictable = entries
      .map((c, index) => ({ c, index }))
      .filter(({ c, index }) => !isPreserved(c, index));

    let evicted = 0;
    for (const { c } of evictable) {
      const remaining = entries.length - evicted;
      const tokens = totalTokens(this.context.list());
      if (remaining <= this.compaction.maxParts && tokens <= this.compaction.maxTokens) break;
      this.context.remove(c.ref.id);
      evicted++;
    }
    return evicted;
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
  async commit(mode: CommitMode, opts: { signal?: AbortSignal } = {}): Promise<void> {
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
      ...(opts.signal ? { signal: opts.signal } : {}),
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
    changeId: ChangeId,
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

  get sessionId(): SessionId | undefined {
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

/**
 * The estimated token cost of one resident context part, for compaction budgeting. Only `text`
 * values carry inline body; named references (`indexed-document`/`drive-document`/`person`) are
 * cheap handles whose cost is negligible — count them as a small fixed overhead.
 */
function tokensOf(ctx: ResolvedContext): number {
  return ctx.value.as === 'text' ? estimateTokens(ctx.value.text) : 1;
}
