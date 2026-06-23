import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ChangeId,
  ContextKind,
  ContextRef,
  ProvenancePayload,
  SourceRef,
  SseEvent,
} from '@ge/contracts';
import { asChangeId } from '@ge/contracts';
import type { CommandLoopEvent, RunCommandsOptions } from '@ge/runtime';
import type { HostEvent } from '@ge/triggers';
import { ProvenanceStore, type ChangeRecord } from './provenance-store.js';
import { renderCommandLine } from './render-command.js';

/**
 * One actuation in a composed plan (ADR-0005 §3 Planner/Executor). The runtime dry-runs the turn
 * (reads + pure transforms, no writes), computes the full effect-set, and emits it for ONE
 * plan-level approval before any effect actuates. `command` is the verbatim CLI line
 * (`set Sales!F2 =SUM(C2:C7)`) the preview renders; `request` is the typed `ActuationRequest` that
 * will execute — they are the SAME object, so what the user sees is exactly what runs.
 *
 * Declared structurally here (mirroring the fixed runtime contract) so the web-shell stays
 * independently verifiable while the runtime half is built in parallel.
 */
export interface PlanEffect {
  request: ActuationRequest;
  /** The verbatim CLI line shown on the plan-approval card. */
  command: string;
  /**
   * Optional dry-run preview the runtime computes while planning (reads + pure transforms, no
   * writes). Purely presentational: the plan-approval card expands an effect to show the resolved
   * value and/or a before→after diff so the reviewer sees what the effect will produce before it
   * runs. Absent for effects the dry-run could not resolve (e.g. a not-yet-readable target).
   */
  dryRun?: EffectDryRun;
}

/**
 * The resolved preview of a single plan effect, produced by the runtime's no-write dry-run. View
 * data only — it never gates or actuates; the plan-approval card renders it to make each effect's
 * outcome legible before the user approves the whole plan.
 */
export interface EffectDryRun {
  /** A short human label for the effect's target, e.g. "Sales!F2" or "“liability cap”". */
  target?: string;
  /** The resolved value the effect will write/insert, e.g. "=C2-D2" or "$184,000". */
  resolved?: string;
  /** The current value at the target before the effect, for a before→after preview. */
  before?: string;
  /** The value the target will hold after the effect, for a before→after preview. */
  after?: string;
}

/**
 * The plan-level approval callback the command loop awaits before executing a composed plan
 * (ADR-0005 §3). **Fail-closed:** with no approver the whole plan is blocked. Supersedes the
 * per-write `approveWrite` for composed plans. Mirrors the fixed runtime contract structurally.
 */
export type ApprovePlan = (effects: PlanEffect[]) => boolean | Promise<boolean>;

/**
 * The `plan-preview` command-loop event (ADR-0005 §3): the runtime emits the dry-run effect-set
 * just before it awaits `approvePlan`. Declared structurally here so the controller can reduce it
 * without depending on the not-yet-shipped runtime `CommandLoopEvent` variant.
 */
export interface PlanPreviewEvent {
  type: 'plan-preview';
  turn: number;
  effects: PlanEffect[];
}

/**
 * `RunCommandsOptions` (from `@ge/runtime`) augmented with the ADR-0005 plan-level approver. The
 * runtime's option type gains `approvePlan` in parallel; we widen it locally so passing the
 * approver type-checks against the current runtime types without coupling to that team's progress.
 */
export type PlanRunCommandsOptions = RunCommandsOptions & { approvePlan?: ApprovePlan };

/** Narrow an arbitrary loop event to the structural `plan-preview` shape. */
function isPlanPreview(ev: { type: string }): ev is PlanPreviewEvent {
  return ev.type === 'plan-preview';
}

/**
 * The subset of `AssistSession` the panel drives. `AssistSession` satisfies this structurally,
 * so the controller is unit-testable against a fake and carries no host/network dependency.
 */
export interface AssistLike {
  readonly context: { size: number };
  attachRef(ref: ContextRef): Promise<void>;
  detach(id: string): void;
  ask(query: string, opts?: { signal?: AbortSignal }): AsyncGenerator<SseEvent>;
  apply(
    kind: ActuationRequest['kind'],
    params: ActuationParams,
    changeId: ChangeId,
  ): Promise<ActuationResult>;
  /**
   * The ADR-0004 read-many/write-one command loop, extended with ADR-0005 plan execution. Streams
   * `SseEvent`s (tokens/citations) plus `CommandLoopEvent`s (loop mechanics, incl. `plan-preview`);
   * calls `opts.approveWrite` for EVERY compiled per-write (fail-closed) and `opts.approvePlan`
   * once for a composed plan's full effect-set (fail-closed) before any effect actuates.
   */
  runCommands(
    task: string,
    opts?: PlanRunCommandsOptions,
  ): AsyncGenerator<SseEvent | CommandLoopEvent>;
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
  /** The turn was cancelled by the user mid-stream (distinct from a stream `error`). */
  cancelled?: boolean;
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
  changeId: ChangeId;
  kind: ActuationRequest['kind'];
  params: ActuationParams;
  label: string;
  status: 'pending' | 'applying' | 'applied' | 'blocked' | 'degraded' | 'failed';
  detail?: string;
  /**
   * Provenance of an applied write, surfaced for the drill-down (agent, identity, sources,
   * timestamp). Presentational only — set when the turn that produced the write carried provenance.
   */
  provenance?: ProvenancePayload;
  /**
   * An optional Excel linked-entity card to render under a `write-cells`/`set-entity-card`
   * proposal (mockup `2-excel.html`'s `◆ Northwind Cloud` card). View data only — enriched from the
   * unit, never stored in the workbook.
   */
  entityCard?: EntityCard;
}

/**
 * A read-only, presentational rendering of an Excel linked-entity card (the mockup's
 * `◆ <entity>` expandable card). The bridge owns the real linked-entity write; this is just the
 * panel's view of the enrichment loaded from the unit, never persisted into the file.
 */
export interface EntityCard {
  title: string;
  subtitle?: string;
  rows: { key: string; value: string }[];
  /** The "loaded from the unit · not stored in the workbook" footnote. */
  footnote?: string;
}

/**
 * An in-session skill (ADR-0005 `def`): a named, parameterized program the user can invoke, which
 * expands into a reviewable plan. View-model only — the controller's `skills` presenter is
 * READ-ONLY: it surfaces what was registered (the `def` confirmation) and lets the view preview a
 * skill's plan via the existing fail-closed plan gate. It carries NO execution/gate logic of its
 * own; invoking a skill still routes through `runCommands` and the plan-approval card.
 */
export interface Skill {
  /** Skill name as registered via `def`, e.g. "flag-vendor-risk". */
  name: string;
  /** A one-line description of what the skill does. */
  description?: string;
  /** The declared parameters, in order, e.g. [{ name: "vendor", example: "Northwind" }]. */
  params: SkillParam[];
  /** True once the `def` registration was confirmed by the runtime (shows the "registered" badge). */
  registered?: boolean;
  /** The verbatim `def` line the runtime echoed back, shown as the registration confirmation. */
  def?: string;
}

export interface SkillParam {
  name: string;
  example?: string;
}

/**
 * A pending write awaiting the user's per-write decision — the fail-closed human-in-the-loop of
 * ADR-0004's command loop. `command` is the verbatim CLI line (`set Sales!F2 =C2-D2`) the approval
 * card renders; the `runCommands` `approveWrite` callback awaits the user's Accept/Reject here.
 */
export interface PendingWrite {
  changeId: ChangeId;
  kind: ActuationRequest['kind'];
  /** The verbatim command line shown on the approval card. */
  command: string;
}

/**
 * A composed plan awaiting the user's ONE plan-level decision (ADR-0005 §3). The runtime has
 * type-checked the plan and dry-run it (reads + pure transforms, no writes) and emitted the full
 * effect-set; nothing actuates until the user calls `approvePlan()`. `effects` are rendered verbatim
 * on the plan card and are exactly the `ActuationRequest`s that execute on approval (no
 * render-benign / execute-malicious divergence). `summary` is the one-line count header
 * (e.g. "3 writes + 2 comments").
 */
export interface PendingPlan {
  effects: PlanEffect[];
  /** One-line count summary of the effect-set, e.g. "3 writes + 2 comments". */
  summary: string;
}

/** One narrated step of the command loop, surfaced so the user can see the loop's progress. */
export interface RunStep {
  id: string;
  /** Mirrors the `CommandLoopEvent.type` (plus a synthetic `error` for run failures). */
  kind:
    | 'turn-start'
    | 'command'
    | 'plan-preview'
    | 'read-result'
    | 'write-result'
    | 'no-fence'
    | 'capped'
    | 'done'
    | 'exhausted'
    | 'error';
  text: string;
}

export interface PanelState {
  messages: ChatMessage[];
  chips: ContextChip[];
  suggestions: Suggestion[];
  proposals: Proposal[];
  changes: ChangeRecord[];
  /** The command-loop transcript (ADR-0004 read-many/write-one steps). */
  steps: RunStep[];
  /** The single write awaiting approval, if the loop is currently gated on the user (ADR-0004). */
  pendingWrite?: PendingWrite;
  /** The composed plan awaiting ONE plan-level approval, if the loop is gated on it (ADR-0005). */
  pendingPlan?: PendingPlan;
  /**
   * In-session skills (ADR-0005 `def`) registered for this surface, surfaced so the user can see
   * what's invokable and preview a skill's plan. Optional/back-compat: absent means "no skills
   * registered". READ-ONLY in the controller — populated by `registerSkills`, never gated here.
   */
  skills?: Skill[];
  busy: boolean;
  error?: string;
}

const EMPTY_STATE: PanelState = {
  messages: [],
  chips: [],
  suggestions: [],
  proposals: [],
  changes: [],
  steps: [],
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
  /** Aborts the in-flight turn's network/stream; cleared when the turn settles. */
  private inflight: AbortController | undefined;
  /**
   * Resolves the `approveWrite` promise the command loop is awaiting. Set while a `pendingWrite` is
   * staged; `approvePendingWrite()`/`rejectPendingWrite()` call it. Fail-closed: if the loop is
   * abandoned (cancel/teardown) without a decision, we resolve `false` so no write actuates.
   */
  private resolvePendingWrite: ((approved: boolean) => void) | undefined;
  /**
   * Resolves the `approvePlan` promise the command loop is awaiting for a composed plan (ADR-0005).
   * Set while a `pendingPlan` is staged; `approvePlan()`/`rejectPlan()` call it. Fail-closed: if the
   * plan is abandoned (cancel/teardown/error) without a decision, we resolve `false` so the WHOLE
   * plan is blocked and no effect actuates.
   */
  private resolvePendingPlan: ((approved: boolean) => void) | undefined;

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

    const controller = new AbortController();
    this.inflight = controller;
    const sources: SourceRef[] = [];
    try {
      for await (const ev of this.session.ask(q, { signal: controller.signal })) {
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
      // A user cancellation surfaces as an AbortError (or the signal flips aborted); mark it as a
      // cancelled turn rather than a red error — the partial answer stays, just no longer streaming.
      if (controller.signal.aborted || isAbortError(err)) {
        this.patchMessage(reply.id, () => ({ cancelled: true }));
        // A cancelled turn never fully landed, so it carries no provenance worth stamping onto a
        // later write. Provenance arrives only at end-of-stream today, so this is belt-and-braces:
        // drop any value from this turn so applyProposal can't stamp a write with a half-landed turn.
        this.lastProvenance = undefined;
      } else {
        this.patchMessage(reply.id, () => ({ error: errorText(err) }));
      }
    } finally {
      // Clear the stored controller so a later cancel() after settle is a clean no-op (only if it is
      // still ours — a queued turn that already replaced it must keep its own controller).
      if (this.inflight === controller) this.inflight = undefined;
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

  /**
   * Cancel the in-flight turn's network/stream. No-op when idle. A queued turn (pendingQuery) still
   * drains afterwards — cancelling the current ask should run the one the user lined up behind it,
   * not discard it. Office.js host writes already under way are not aborted (not abortable); this
   * targets the assist stream only.
   */
  cancel(): void {
    this.inflight?.abort();
    // A loop gated on the user when cancelled must release fail-closed, so its awaited approval
    // resolves `false` and nothing actuates after the abort — for both the per-write gate and the
    // ADR-0005 plan-level gate.
    this.settlePendingWrite(false);
    this.settlePendingPlan(false);
  }

  // ---- command loop (ADR-0004) --------------------------------------------

  /**
   * Drive the ADR-0004 read-many/write-one command loop with UI-backed, per-write approval. Streams
   * the grounded answer (tokens/citations) into an assistant message exactly like `send()`, while
   * narrating the loop's mechanics as `steps`. EVERY model-emitted write is staged as a
   * `pendingWrite` and actuates only after the user calls `approvePendingWrite()` — the fail-closed
   * human-in-the-loop. `send()`/`ask()` stay intact and untouched.
   */
  async runCommands(task: string): Promise<void> {
    const t = task.trim();
    if (!t) return;
    if (this.state.busy) {
      this.pendingQuery = t;
      return;
    }

    const userMsg: ChatMessage = { id: this.id('u'), role: 'user', text: t };
    const reply: ChatMessage = { id: this.id('a'), role: 'assistant', text: '', streaming: true };
    this.set({
      messages: [...this.state.messages, userMsg, reply],
      busy: true,
      error: undefined,
      suggestions: [],
      steps: [],
    });

    const controller = new AbortController();
    this.inflight = controller;
    const sources: SourceRef[] = [];

    // The per-write approver (ADR-0004): stage the compiled request and await the user's decision.
    const approveWrite = (request: ActuationRequest): Promise<boolean> =>
      new Promise<boolean>((resolve) => {
        // Defensive: if a prior decision is somehow still open, release it false first.
        this.settlePendingWrite(false);
        this.resolvePendingWrite = resolve;
        this.set({
          pendingWrite: {
            changeId: request.changeId,
            kind: request.kind,
            command: renderCommandLine(request),
          },
        });
      });

    // The plan-level approver (ADR-0005): stage the full dry-run effect-set and await ONE decision.
    // Fail-closed: nothing here resolves `true` except an explicit `approvePlan()`.
    const approvePlan: ApprovePlan = (effects) =>
      new Promise<boolean>((resolve) => {
        this.settlePendingPlan(false);
        this.resolvePendingPlan = resolve;
        this.set({ pendingPlan: { effects, summary: summarizeEffects(effects) } });
      });

    const opts: PlanRunCommandsOptions = {
      signal: controller.signal,
      approveWrite,
      approvePlan,
    };

    try {
      for await (const ev of this.session.runCommands(t, opts)) {
        // The plan-preview variant is not (yet) in the runtime `CommandLoopEvent` union; narrow it
        // structurally so we reduce it without coupling to the parallel runtime build.
        if (isPlanPreview(ev)) {
          this.addStep('plan-preview', summarizeEffects(ev.effects));
          continue;
        }
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
            // `SseEvent` error (stream-level), distinct from a CommandLoopEvent.
            this.patchMessage(reply.id, () => ({ error: ev.message }));
            this.addStep('error', ev.message);
            break;
          case 'turn-start':
            this.addStep('turn-start', `Turn ${ev.turn}`);
            break;
          case 'command':
            this.addStep('command', commandStepText(ev));
            break;
          case 'read-result':
            this.addStep('read-result', `read ${ev.intentLabel}`);
            break;
          case 'write-result':
            this.addStep('write-result', writeStepText(ev));
            // The decision has been consumed by the loop; clear the staged pending write.
            this.clearPendingWrite();
            break;
          case 'no-fence':
            this.addStep('no-fence', `Turn ${ev.turn}: no command block — re-prompting`);
            break;
          case 'capped':
            this.addStep('capped', ev.reason);
            break;
          case 'done':
            this.addStep('done', 'Done');
            break;
          case 'exhausted':
            this.addStep('exhausted', `Stopped after ${ev.turns} turns`);
            break;
          default:
            break;
        }
      }
    } catch (err) {
      if (controller.signal.aborted || isAbortError(err)) {
        this.patchMessage(reply.id, () => ({ cancelled: true }));
        this.lastProvenance = undefined;
      } else {
        this.patchMessage(reply.id, () => ({ error: errorText(err) }));
        this.addStep('error', errorText(err));
      }
    } finally {
      // Any write or plan still gated when the loop ends releases fail-closed (never default-accept).
      this.settlePendingWrite(false);
      this.settlePendingPlan(false);
      if (this.inflight === controller) this.inflight = undefined;
      this.patchMessage(reply.id, () => ({ streaming: false }));
      this.set({ busy: false, changes: this.store.list() });
      const next = this.pendingQuery;
      if (next !== undefined) {
        this.pendingQuery = undefined;
        void this.send(next);
      }
    }
  }

  /** Approve the staged write — resolves the loop's `approveWrite` with `true` so it actuates. */
  approvePendingWrite(): void {
    this.settlePendingWrite(true);
  }

  /** Reject the staged write — resolves the loop's `approveWrite` with `false`; nothing actuates. */
  rejectPendingWrite(): void {
    this.settlePendingWrite(false);
  }

  /** Resolve the awaited decision (if any) and stop showing the approval card. */
  private settlePendingWrite(approved: boolean): void {
    const resolve = this.resolvePendingWrite;
    if (!resolve) return;
    this.resolvePendingWrite = undefined;
    // On reject, drop the card now; on approve, keep it until the `write-result` narrates outcome.
    if (!approved) this.clearPendingWrite();
    resolve(approved);
  }

  private clearPendingWrite(): void {
    if (this.state.pendingWrite) this.set({ pendingWrite: undefined });
  }

  /**
   * Approve the staged composed plan (ADR-0005) — resolves the loop's `approvePlan` with `true`, so
   * the executor runs the previewed effect-set (each effect still gated/provenanced one-by-one).
   */
  approvePlan(): void {
    this.settlePendingPlan(true);
  }

  /** Reject the staged plan — resolves `approvePlan` with `false`; the WHOLE plan is blocked. */
  rejectPlan(): void {
    this.settlePendingPlan(false);
  }

  /** Resolve the awaited plan decision (if any) and stop showing the plan-approval card. */
  private settlePendingPlan(approved: boolean): void {
    const resolve = this.resolvePendingPlan;
    if (!resolve) return;
    this.resolvePendingPlan = undefined;
    // Drop the card on a decision either way: the executor's per-effect outcomes narrate as steps.
    this.clearPendingPlan();
    resolve(approved);
  }

  private clearPendingPlan(): void {
    if (this.state.pendingPlan) this.set({ pendingPlan: undefined });
  }

  private addStep(kind: RunStep['kind'], text: string): void {
    this.set({ steps: [...this.state.steps, { id: this.id('step'), kind, text }] });
  }

  // ---- actuation review ---------------------------------------------------

  /** Stage a reviewable, reversible change for the user to confirm. */
  propose(kind: ActuationRequest['kind'], params: ActuationParams, label: string): Proposal {
    const proposal: Proposal = {
      changeId: asChangeId(this.id('c')),
      kind,
      params,
      label,
      status: 'pending',
    };
    this.set({ proposals: [...this.state.proposals, proposal] });
    return proposal;
  }

  /** Apply a staged proposal through the session (gate → bridge), recording the outcome. */
  async applyProposal(changeId: ChangeId): Promise<void> {
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

  // ---- skills (ADR-0005 `def`) — READ-ONLY presenters ---------------------

  /**
   * Surface the in-session skills registered for this surface (the `def` confirmations). Purely a
   * view presenter: it records what is invokable so the panel can list it and preview a skill's
   * plan. It adds NO execution or gate logic — invoking a skill still goes through `runCommands` and
   * the fail-closed plan-approval card. Replaces the skills slice wholesale (latest registration
   * wins), mirroring how the runtime re-emits the full `def` set.
   */
  registerSkills(skills: Skill[]): void {
    this.set({ skills });
  }

  /** The currently-registered skills (read-only snapshot). */
  listSkills(): Skill[] {
    return this.state.skills ?? [];
  }

  /**
   * Invoke a registered skill with bound argument values. This is a thin presenter over the EXISTING
   * agentic command loop: it composes the skill call as a task line and routes it through
   * `runCommands`, so the skill's plan still lands on the fail-closed plan-approval card and nothing
   * actuates without explicit approval. It introduces no new gate of its own. No-op if the named
   * skill is not registered.
   */
  async invokeSkill(name: string, args: Record<string, string> = {}): Promise<void> {
    const skill = (this.state.skills ?? []).find((s) => s.name === name);
    if (!skill) return;
    await this.runCommands(renderSkillCall(skill, args));
  }

  // ---- internals ----------------------------------------------------------

  private set(patch: Partial<PanelState>): void {
    this.state = { ...this.state, ...patch };
    for (const listener of this.listeners) listener(this.state);
  }

  private setChip(id: string, patch: Partial<ContextChip>): void {
    this.set({ chips: this.state.chips.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  }

  private setProposal(changeId: ChangeId, patch: Partial<Proposal>): void {
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

/**
 * Render a skill invocation as the verbatim task line the agentic loop receives, e.g.
 * `flag-vendor-risk vendor="Northwind" tier="1"`. Pure/total — params with no bound value fall back
 * to their declared example, then to an empty quoted operand, so the call is always well-formed.
 */
function renderSkillCall(skill: Skill, args: Record<string, string>): string {
  const parts = skill.params.map((p) => {
    const value = args[p.name] ?? p.example ?? '';
    return `${p.name}="${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
  });
  return [skill.name, ...parts].join(' ');
}

/**
 * A one-line count summary of a plan's effect-set for the preview header, e.g. "3 writes + 2
 * comments". Groups by a human label for each `ActuationRequest.kind` and pluralizes. Pure/total —
 * an empty set degrades to "no effects" rather than rendering an empty header.
 */
function summarizeEffects(effects: readonly PlanEffect[]): string {
  if (effects.length === 0) return 'no effects';
  const counts = new Map<string, number>();
  const order: string[] = [];
  for (const e of effects) {
    const label = effectNoun(e.request.kind);
    if (!counts.has(label)) order.push(label);
    counts.set(label, (counts.get(label) ?? 0) + 1);
  }
  return order
    .map((label) => {
      const n = counts.get(label) ?? 0;
      return `${n} ${label}${n === 1 ? '' : 's'}`;
    })
    .join(' + ');
}

/** A human noun for an actuation kind, used by the plan summary header. */
function effectNoun(kind: ActuationRequest['kind']): string {
  switch (kind) {
    case 'write-cells':
    case 'tracked-change':
      return 'write';
    case 'add-comment':
    case 'comment-reply':
      return 'comment';
    case 'format-cells':
      return 'format';
    default:
      return kind;
  }
}

/** A one-line label for a `command` loop step: the parsed verb (or the corrective parse error). */
function commandStepText(ev: Extract<CommandLoopEvent, { type: 'command' }>): string {
  if ('error' in ev.compiled) return `error: ${ev.compiled.error}`;
  return ev.command.verb;
}

/** A one-line label for a `write-result` loop step: the write kind + its outcome. */
function writeStepText(ev: Extract<CommandLoopEvent, { type: 'write-result' }>): string {
  const r = ev.result;
  const outcome = r.ok ? (r.degraded ? 'degraded' : 'applied') : (r.error?.code ?? 'failed');
  // Observability: the change landed but its durable provenance dropped — make it visible, not silent.
  const provenance = r.ok && r.provenanceDropped ? ' (⚠ provenance not recorded)' : '';
  return `${r.kind} — ${outcome}${provenance}`;
}

/** A fetch aborted via AbortSignal rejects with a DOMException/Error named 'AbortError'. */
function isAbortError(err: unknown): boolean {
  return err instanceof Error && err.name === 'AbortError';
}
