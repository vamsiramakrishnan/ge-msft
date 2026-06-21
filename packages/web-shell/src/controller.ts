import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ContextKind,
  ContextRef,
  ProvenancePayload,
  SourceRef,
  SseEvent,
} from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { ProvenanceStore, type ChangeRecord } from './provenance-store.js';

/**
 * The subset of `AssistSession` the panel drives. `AssistSession` satisfies this structurally,
 * so the controller is unit-testable against a fake and carries no host/network dependency.
 */
export interface AssistLike {
  readonly context: { size: number };
  attachRef(ref: ContextRef): Promise<void>;
  detach(id: string): void;
  ask(query: string): AsyncGenerator<SseEvent>;
  apply(
    kind: ActuationRequest['kind'],
    params: ActuationParams,
    changeId: string,
  ): Promise<ActuationResult>;
  ingest(event: HostEvent): Promise<void>;
  readonly sessionId?: string;
}

/** Lists what can be attached right now (the bridge). */
export interface ContextLister {
  listContext(): Promise<ContextRef[]>;
}

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant';
  text: string;
  streaming?: boolean;
  sources?: SourceRef[];
  error?: string;
}

export interface ContextChip {
  id: string;
  title: string;
  kind: ContextKind;
  attached: boolean;
  preview?: string;
}

export interface Suggestion {
  id: string;
  title: string;
  detail?: string;
  query?: string;
}

export interface Proposal {
  changeId: string;
  kind: ActuationRequest['kind'];
  params: ActuationParams;
  label: string;
  status: 'pending' | 'applying' | 'applied' | 'blocked' | 'degraded' | 'failed';
  detail?: string;
}

export interface PanelState {
  messages: ChatMessage[];
  chips: ContextChip[];
  suggestions: Suggestion[];
  proposals: Proposal[];
  changes: ChangeRecord[];
  busy: boolean;
  error?: string;
}

const EMPTY_STATE: PanelState = {
  messages: [],
  chips: [],
  suggestions: [],
  proposals: [],
  changes: [],
  busy: false,
};

/**
 * The surface-agnostic panel logic — the brain behind the task pane. It drives the assist loop
 * (attach context → ask → stream → review/apply), keeps the immutable `PanelState` a thin React
 * view renders, and exposes `onContext`/`onSuggest`/`onAutomate` so the event Orchestrator feeds
 * straight in. No React, no Office.js here: those live one layer up.
 */
export class PanelController {
  private state: PanelState = EMPTY_STATE;
  private readonly listeners = new Set<(state: PanelState) => void>();
  private readonly refs = new Map<string, ContextRef>();
  private lastProvenance: ProvenancePayload | undefined;
  private seq = 0;
  /** Single-slot queue (latest wins): a turn requested while another is streaming. */
  private pendingQuery: string | undefined;

  constructor(
    private readonly session: AssistLike,
    private readonly lister: ContextLister,
    private readonly store: ProvenanceStore = new ProvenanceStore(),
  ) {}

  getState(): PanelState {
    return this.state;
  }

  subscribe(listener: (state: PanelState) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  // ---- context tray -------------------------------------------------------

  /** Load the attachable-context chips (preserving which are already attached). */
  async refreshContext(): Promise<void> {
    try {
      const attachedIds = new Set(this.state.chips.filter((c) => c.attached).map((c) => c.id));
      const refs = await this.lister.listContext();
      this.refs.clear();
      const chips = refs.map((r): ContextChip => {
        this.refs.set(r.id, r);
        return {
          id: r.id,
          title: r.title,
          kind: r.kind,
          attached: attachedIds.has(r.id),
          ...(r.preview ? { preview: r.preview } : {}),
        };
      });
      this.set({ chips });
    } catch (err) {
      this.set({ error: errorText(err) });
    }
  }

  async attach(id: string): Promise<void> {
    const ref = this.refs.get(id);
    if (!ref) return;
    try {
      await this.session.attachRef(ref);
      this.setChip(id, { attached: true });
    } catch (err) {
      this.set({ error: errorText(err) });
    }
  }

  detach(id: string): void {
    this.session.detach(id);
    this.setChip(id, { attached: false });
  }

  // ---- ask / stream -------------------------------------------------------

  /** Ask a grounded question; stream tokens + citations into the assistant message. */
  async send(query: string): Promise<void> {
    const q = query.trim();
    if (!q) return;
    // A turn requested mid-stream is held (latest wins) and drained when the current turn ends,
    // so an opt-in automated turn is never silently dropped. See `finally` below.
    if (this.state.busy) {
      this.pendingQuery = q;
      return;
    }

    const userMsg: ChatMessage = { id: this.id('u'), role: 'user', text: q };
    const reply: ChatMessage = { id: this.id('a'), role: 'assistant', text: '', streaming: true };
    this.set({
      messages: [...this.state.messages, userMsg, reply],
      busy: true,
      error: undefined,
      suggestions: [],
    });

    const sources: SourceRef[] = [];
    try {
      for await (const ev of this.session.ask(q)) {
        switch (ev.type) {
          case 'token':
            this.patchMessage(reply.id, (m) => ({ text: m.text + ev.text }));
            break;
          case 'citation':
            sources.push(ev.source);
            this.patchMessage(reply.id, () => ({ sources: [...sources] }));
            break;
          case 'provenance':
            this.lastProvenance = ev.payload;
            break;
          case 'error':
            this.patchMessage(reply.id, () => ({ error: ev.message }));
            break;
          default:
            break;
        }
      }
    } catch (err) {
      this.patchMessage(reply.id, () => ({ error: errorText(err) }));
    } finally {
      this.patchMessage(reply.id, () => ({ streaming: false }));
      this.set({ busy: false });
      // Drain a queued turn, if any. Clearing the slot first avoids re-enqueueing it; the call
      // is scheduled (not awaited) so draining does not re-enter synchronously while busy.
      const next = this.pendingQuery;
      if (next !== undefined) {
        this.pendingQuery = undefined;
        void this.send(next);
      }
    }
  }

  // ---- actuation review ---------------------------------------------------

  /** Stage a reviewable, reversible change for the user to confirm. */
  propose(kind: ActuationRequest['kind'], params: ActuationParams, label: string): Proposal {
    const proposal: Proposal = { changeId: this.id('c'), kind, params, label, status: 'pending' };
    this.set({ proposals: [...this.state.proposals, proposal] });
    return proposal;
  }

  /** Apply a staged proposal through the session (gate → bridge), recording the outcome. */
  async applyProposal(changeId: string): Promise<void> {
    const proposal = this.state.proposals.find((p) => p.changeId === changeId);
    if (!proposal || proposal.status !== 'pending') return;

    // Flip status synchronously before awaiting so a second (e.g. double-click) call bails at
    // the guard above — preventing a double host write / duplicate ChangeRecord.
    this.setProposal(changeId, { status: 'applying' });

    const result = await this.session.apply(proposal.kind, proposal.params, proposal.changeId);
    this.store.record(result, this.lastProvenance);

    const status: Proposal['status'] = result.ok
      ? 'applied'
      : result.degraded
        ? 'degraded'
        : result.error?.code === 'blocked'
          ? 'blocked'
          : 'failed';
    this.setProposal(changeId, {
      status,
      ...(result.error ? { detail: result.error.message } : {}),
    });
    this.set({ changes: this.store.list() });
  }

  // ---- event-driven inputs (wire to the Orchestrator) ---------------------

  /** Every host event constructs working context (no model call). */
  readonly onContext = (event: HostEvent): void => {
    void this.session.ingest(event);
  };

  /** A rare, ignorable ambient suggestion chip. */
  readonly onSuggest = (s: { title: string; detail?: string; query?: string }): void => {
    const suggestion: Suggestion = {
      id: this.id('s'),
      title: s.title,
      ...(s.detail ? { detail: s.detail } : {}),
      ...(s.query ? { query: s.query } : {}),
    };
    this.set({ suggestions: [...this.state.suggestions, suggestion] });
  };

  /** An opt-in automated turn (e.g. accepting a suggestion). */
  readonly onAutomate = (query: string): void => {
    void this.send(query);
  };

  dismissSuggestion(id: string): void {
    this.set({ suggestions: this.state.suggestions.filter((s) => s.id !== id) });
  }

  // ---- internals ----------------------------------------------------------

  private set(patch: Partial<PanelState>): void {
    this.state = { ...this.state, ...patch };
    for (const listener of this.listeners) listener(this.state);
  }

  private setChip(id: string, patch: Partial<ContextChip>): void {
    this.set({ chips: this.state.chips.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  }

  private setProposal(changeId: string, patch: Partial<Proposal>): void {
    this.set({
      proposals: this.state.proposals.map((p) =>
        p.changeId === changeId ? { ...p, ...patch } : p,
      ),
    });
  }

  private patchMessage(id: string, patch: (m: ChatMessage) => Partial<ChatMessage>): void {
    this.set({
      messages: this.state.messages.map((m) => (m.id === id ? { ...m, ...patch(m) } : m)),
    });
  }

  private id(prefix: string): string {
    return `${prefix}-${++this.seq}`;
  }
}

function errorText(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
