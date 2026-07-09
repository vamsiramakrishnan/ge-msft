import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ChangeId,
  CommandPlan,
  ContextKind,
  ContextRef,
  ProvenancePayload,
  SourceRef,
  Surface,
  SseEvent,
} from '@ge/contracts';
import { asChangeId, deriveOutput, extractCommandBlock } from '@ge/contracts';
import type {
  AgentView,
  ConversationSummary,
  EngineDataStore,
  ResolvedGrounding,
} from '@ge/gemini-client';
import {
  type CommandLoopEvent,
  type PlanEffect as RuntimePlanEffect,
  type RunCommandsOptions,
  type WorkspaceArtifactSummary,
  type WorkspaceResult,
} from '@ge/runtime';
import type { HostEvent } from '@ge/triggers';
import { ProvenanceStore, type ChangeRecord } from './provenance-store.js';
import { renderCommandLine } from './render-command.js';
import { assertNever } from './assert-never.js';
import { TurnQueue } from './turn-queue.js';
import { ApprovalCoordinator } from './approval-coordinator.js';
import { findContextRefForLocation, synthesizeLocationRef } from './host-location.js';

/**
 * One actuation in a composed plan (ADR-0005 §3 Planner/Executor) — the CANONICAL runtime
 * {@link RuntimePlanEffect} (Finding #6: the controller no longer redeclares it), widened only with
 * the controller-owned PRESENTATIONAL {@link EffectDryRun}. `request` is the typed `ActuationRequest`
 * that executes; `command` is the verbatim CLI line the preview renders — the SAME object, so what
 * the user sees is exactly what runs. The richer `dryRun` (before→after) is a panel enrichment, not
 * part of the execution contract.
 */
export type PlanEffect = RuntimePlanEffect & { dryRun?: EffectDryRun };

/**
 * The resolved preview of a single plan effect, produced by the runtime's no-write dry-run, widened
 * with the panel's before→after diff. View data only — it never gates or actuates; the plan-approval
 * card renders it to make each effect's outcome legible before the user approves the whole plan. The
 * runtime contract carries only `{ target, resolved }`; `before`/`after` are a presentational
 * enrichment the panel computes, so this is a superset (assignable to/from the runtime's `dryRun`).
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

/** The plan-level approver type — the CANONICAL runtime option (Finding #6: not redeclared here). */
export type ApprovePlan = NonNullable<RunCommandsOptions['approvePlan']>;
/** The per-write approver type — the CANONICAL runtime option (Finding #6: not redeclared here). */
type ApproveWrite = NonNullable<RunCommandsOptions['approveWrite']>;

/** Direct pasted CLI is user-authored and fully previewed, so it can stage larger exact programs. */
const DIRECT_COMMAND_LIMIT = 128;

/**
 * The `plan-preview` command-loop event (ADR-0005 §3) — re-exported as the CANONICAL runtime variant
 * (Finding #6: the controller no longer declares a parallel structural copy). The runtime emits the
 * dry-run effect-set just before it awaits `approvePlan`.
 */
export type PlanPreviewEvent = Extract<CommandLoopEvent, { type: 'plan-preview' }>;

/**
 * Back-compat alias for the command-loop options the controller passes to the session. It is now the
 * CANONICAL {@link RunCommandsOptions} verbatim — `approvePlan`/`approveWrite`/`grounding` all live on
 * the runtime contract, so no local widening is needed (Finding #6: the parallel widening is gone).
 */
export type PlanRunCommandsOptions = RunCommandsOptions;

/**
 * The subset of `AssistSession` the panel drives. `AssistSession` satisfies this structurally,
 * so the controller is unit-testable against a fake and carries no host/network dependency.
 */
export interface AssistLike {
  readonly context: { size: number };
  attachRef(ref: ContextRef): Promise<void>;
  detach(id: string): void;
  /**
   * Ask a grounded question. `opts.grounding` is the STRUCTURED resolution of the turn's typed
   * `@`-mentions (Finding #2/#B-wire) — addressed query parts / data stores / files, NOT free-text.
   */
  ask(
    query: string,
    opts?: { signal?: AbortSignal; grounding?: ResolvedGrounding },
  ): AsyncGenerator<SseEvent>;
  /**
   * Apply a staged proposal. Provenance (Finding #4) is passed EXPLICITLY — the provenance captured
   * from the very turn that produced this change — so a later, provenance-less turn can never have
   * its write inherit an earlier turn's leftover provenance.
   */
  apply(
    kind: ActuationRequest['kind'],
    params: ActuationParams,
    changeId: ChangeId,
    provenance?: ProvenancePayload,
  ): Promise<ActuationResult>;
  /**
   * The ADR-0004 read-many/write-one command loop, extended with ADR-0005 plan execution. Streams
   * `SseEvent`s (tokens/citations) plus `CommandLoopEvent`s (loop mechanics, incl. `plan-preview`);
   * calls `opts.approveWrite` for EVERY compiled per-write (fail-closed) and `opts.approvePlan`
   * once for a composed plan's full effect-set (fail-closed) before any effect actuates.
   */
  runCommands(task: string, opts?: RunCommandsOptions): AsyncGenerator<SseEvent | CommandLoopEvent>;
  /**
   * Execute explicit user-authored CLI lines without asking the model to restate them. Implemented by
   * AssistSession; optional so older unit-test fakes can still exercise the controller fallback.
   */
  runCommandProgram?(
    program: string,
    opts?: RunCommandsOptions,
  ): AsyncGenerator<SseEvent | CommandLoopEvent>;
  listConversations?(opts?: {
    pageSize?: number;
    pageToken?: string;
    signal?: AbortSignal;
  }): Promise<{ conversations: ConversationSummary[]; nextPageToken?: string }>;
  resumeSession?(sessionIdOrName: string): void;
  /**
   * The planner pre-stage (EXPERIENCE.md §F): stream one turn that proposes a confirmable
   * {@link CommandPlan} for a complex free-text request, WITHOUT reading or writing the document.
   */
  plan(
    task: string,
    opts?: { signal?: AbortSignal; grounding?: ResolvedGrounding },
  ): Promise<{ plan: CommandPlan | null; errors: string[]; needsClarification: boolean }>;
  ingest(event: HostEvent): Promise<void>;
  readonly sessionId?: string;
}

/** Lists what can be attached right now (the bridge). */
export interface ContextLister {
  listContext(): Promise<ContextRef[]>;
  canRevealContext?(ref: ContextRef): boolean;
  revealContext?(ref: ContextRef): Promise<void>;
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
  revealable?: boolean;
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

export interface ConversationItem {
  name: string;
  id: string;
  title: string;
  turnCount: number;
  isPinned: boolean;
  active: boolean;
  state?: string;
  startedAt?: string;
  endedAt?: string;
  updatedAt?: string;
}

export interface ConversationsState {
  items: ConversationItem[];
  loading: boolean;
  loaded: boolean;
  error?: string;
  nextPageToken?: string;
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

/**
 * The planner's {@link CommandPlan} awaiting the user's confirm BEFORE the executor runs
 * (EXPERIENCE.md §F — the front-door stage for complex free-text). This is the high-level INTENTION
 * (intent · scope · ordered steps · exclusions), distinct from {@link PendingPlan} (the executor's
 * dry-run effect-set). On confirm the controller runs the normalized, user-confirmed plan through
 * `runCommands` — which then stages its own `pendingPlan` for the effect-level gate. A plan carrying
 * `clarify` lines is surfaced as a question instead (it never reaches confirm).
 */
export interface PendingCommandPlan {
  plan: CommandPlan;
  /** The original free-text task, preserved for display and as context inside the confirmed plan. */
  task: string;
  grounding?: ResolvedGrounding;
}

export interface PendingPlanClarification {
  /** The planner request that produced the clarify question. */
  task: string;
  questions: string[];
  grounding?: ResolvedGrounding;
}

function renderConfirmedPlanTask(pending: PendingCommandPlan): string {
  const { plan, task } = pending;
  const lines: string[] = [
    'Execute this user-confirmed plan in the open Microsoft 365 surface.',
    'Treat the plan as approved intent only: read live host content before any write, respect exclusions, emit only the supported cmd protocol, and let the normal preview/approval gate run.',
    '',
    '<confirmed_plan>',
    `original_request: ${task}`,
    `intent: ${plan.intent}`,
    `surface: ${plan.surface}`,
  ];
  if (plan.scope) {
    lines.push(
      `scope: ${plan.scope.ref ? `${plan.scope.kind} ${plan.scope.ref}` : plan.scope.kind}`,
    );
  }
  for (const ground of plan.ground) {
    lines.push(`ground: ${ground.ref ? `${ground.kind} ${ground.ref}` : ground.kind}`);
  }
  for (const hint of plan.context) lines.push(`context: ${hint}`);
  for (const [i, step] of plan.steps.entries()) lines.push(`step ${i + 1}: ${step}`);
  for (const exclude of plan.excludes) lines.push(`exclude: ${exclude}`);
  if (plan.confidence) lines.push(`confidence: ${plan.confidence}`);
  lines.push('</confirmed_plan>');
  return lines.join('\n');
}

function renderConfirmedPlanDisplayText(pending: PendingCommandPlan): string {
  const { plan } = pending;
  const scope = plan.scope
    ? plan.scope.ref
      ? `${plan.scope.kind} ${plan.scope.ref}`
      : plan.scope.kind
    : 'current context';
  const stepCount = plan.steps.length === 1 ? '1 step' : `${plan.steps.length} steps`;
  return `/execute approved ${plan.intent} plan · ${plan.surface} · ${scope} · ${stepCount}`;
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
    | 'code-execution'
    | 'activity'
    | 'error';
  text: string;
  artifact?: RunStepArtifact;
}

export interface RunStepArtifact {
  title: string;
  meta: string[];
  preview?: string;
  matches?: Array<{ line: number; text: string }>;
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
   * The planner's high-level {@link CommandPlan} awaiting the user's confirm before the executor runs
   * (EXPERIENCE.md §F — the complex-free-text front door). Distinct from `pendingPlan`.
   */
  pendingCommandPlan?: PendingCommandPlan;
  /** A planner clarify question awaiting the user's natural-language answer. */
  pendingPlanClarification?: PendingPlanClarification;
  /**
   * In-session skills (ADR-0005 `def`) registered for this surface, surfaced so the user can see
   * what's invokable and preview a skill's plan. Optional/back-compat: absent means "no skills
   * registered". READ-ONLY in the controller — populated by `registerSkills`, never gated here.
   */
  skills?: Skill[];
  /**
   * Skills (agents) discovered via `:listAvailableAgentViews` at boot (Task 6). Populated once by
   * `setDiscoveredCatalog`; a future `@`-picker UI reads this to offer agent mentions. Empty until
   * discovery completes or when discovery is disabled/unavailable.
   */
  availableAgents: AgentView[];
  /**
   * Federated data stores discovered from the engine's config at boot (Task 6). Populated once by
   * `setDiscoveredCatalog`; a future `@`-picker UI reads this to offer data-store mentions. Empty
   * until discovery completes or when discovery is disabled/unavailable.
   */
  availableDataStores: EngineDataStore[];
  conversations: ConversationsState;
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
  availableAgents: [],
  availableDataStores: [],
  conversations: { items: [], loading: false, loaded: false },
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
  private seq = 0;
  /**
   * Finding #4: provenance is TURN-SCOPED, never ambient. This holds ONLY the provenance of the
   * CURRENTLY-streaming turn — captured from that turn's `provenance` SSE event, read by `propose()`
   * to stamp the proposal it creates, and CLEARED on every turn boundary (start, and every
   * success/error/cancel settle). A proposal created by a provenance-less turn therefore carries no
   * provenance, and `applyProposal` stamps the proposal's OWN captured provenance — never a leftover.
   */
  private currentTurnProvenance: ProvenancePayload | undefined;
  /** Aborts the in-flight turn's network/stream; cleared when the turn settles. */
  private inflight: AbortController | undefined;
  /**
   * The single-slot, mode-typed turn queue (E-full): a turn requested while another is streaming,
   * drained back through its OWN route so a queued mode is never downgraded (Finding #3).
   */
  private readonly turnQueue = new TurnQueue();
  /**
   * The fail-closed approval state machine (E-full): owns the per-write changeId + the two resolver
   * promises the loop awaits (Finding #6). The `pendingWrite`/`pendingPlan` VIEW slice stays in
   * `PanelState`, pushed here via the `set` callbacks below.
   */
  private readonly approvals = new ApprovalCoordinator(
    (write) => this.set({ pendingWrite: write }),
    (plan) => this.set({ pendingPlan: plan }),
  );

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
          ...(this.lister.revealContext && (this.lister.canRevealContext?.(r) ?? true)
            ? { revealable: true }
            : {}),
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

  async reveal(id: string): Promise<void> {
    const ref = this.refs.get(id);
    if (!ref || !this.lister.revealContext) return;
    try {
      await this.lister.revealContext(ref);
    } catch (err) {
      this.set({ error: errorText(err) });
    }
  }

  /**
   * Navigation-only reveal for locations surfaced outside the context tray, such as a plan-preview
   * target or an assistant answer that names a cell/range. It reuses the same host bridge as context
   * chips. If the location cannot be represented as a surface-native ref, this fails closed.
   */
  async revealLocation(surface: Surface, location: string): Promise<void> {
    if (!this.lister.revealContext) return;
    const ref =
      findContextRefForLocation(this.refs.values(), surface, location) ??
      synthesizeLocationRef(surface, location);
    if (!ref) return;
    if (this.lister.canRevealContext && !this.lister.canRevealContext(ref)) return;
    try {
      await this.lister.revealContext(ref);
    } catch (err) {
      this.set({ error: errorText(err) });
    }
  }

  // ---- ask / stream -------------------------------------------------------

  /**
   * Ask a grounded question; stream tokens + citations into the assistant message. `grounding` is the
   * structured resolution of the turn's typed `@`-mentions (Finding #2/#B-wire), forwarded to the
   * session as request grounding — never inlined into the prompt string.
   */
  async send(query: string, grounding?: ResolvedGrounding): Promise<void> {
    const q = query.trim();
    if (!q) return;
    // A turn requested mid-stream is held (latest wins) and drained when the current turn ends, so an
    // opt-in automated turn is never silently dropped. Queued AS AN ASK so it drains back through
    // send() (Finding #3) — not collapsed into a string a later drain might mis-route.
    if (this.state.busy) {
      this.turnQueue.enqueue({ mode: 'ask', query: q, ...(grounding ? { grounding } : {}) });
      return;
    }

    const userMsg: ChatMessage = { id: this.id('u'), role: 'user', text: q };
    const reply: ChatMessage = { id: this.id('a'), role: 'assistant', text: '', streaming: true };
    this.beginTurn({ messages: [...this.state.messages, userMsg, reply] });

    const controller = new AbortController();
    this.inflight = controller;
    const sources: SourceRef[] = [];
    let replyText = '';
    let recoverAsCommandTask: string | undefined;
    try {
      for await (const ev of this.session.ask(q, {
        signal: controller.signal,
        ...(grounding ? { grounding } : {}),
      })) {
        switch (ev.type) {
          case 'token':
            replyText += ev.text;
            this.patchMessage(reply.id, (m) => ({ text: m.text + ev.text }));
            break;
          case 'activity':
            this.addStep('activity', ev.text);
            break;
          case 'citation':
            sources.push(ev.source);
            this.patchMessage(reply.id, () => ({ sources: [...sources] }));
            break;
          case 'provenance':
            // Finding #4: capture THIS turn's provenance into the turn-local; `propose()` stamps it
            // onto the proposal it creates. Cleared at the next turn boundary, never left ambient.
            this.currentTurnProvenance = ev.payload;
            break;
          case 'error':
            this.patchMessage(reply.id, () => ({ error: ev.message }));
            break;
          default:
            break;
        }
      }
      if (!controller.signal.aborted && extractCommandBlock(replyText) !== null) {
        recoverAsCommandTask = q;
        this.patchMessage(reply.id, () => ({
          text: 'Detected Office command output in a chat turn. Continuing through the Office command route so the add-in can read/apply changes safely.',
        }));
      }
    } catch (err) {
      // A user cancellation surfaces as an AbortError (or the signal flips aborted); mark it as a
      // cancelled turn rather than a red error — the partial answer stays, just no longer streaming.
      if (controller.signal.aborted || isAbortError(err)) {
        this.patchMessage(reply.id, () => ({ cancelled: true }));
        // A cancelled turn never fully landed, so it carries no provenance worth stamping onto a
        // later write — drop it so applyProposal can't stamp a write with a half-landed turn.
        this.currentTurnProvenance = undefined;
      } else {
        this.patchMessage(reply.id, () => ({ error: errorText(err) }));
      }
    } finally {
      // Clear the stored controller so a later cancel() after settle is a clean no-op (only if it is
      // still ours — a queued turn that already replaced it must keep its own controller).
      if (this.inflight === controller) this.inflight = undefined;
      this.patchMessage(reply.id, () => ({ streaming: false }));
      this.set({ busy: false });
      if (recoverAsCommandTask) {
        await this.runCommands(recoverAsCommandTask, grounding, 'Continue in Office command route');
      } else {
        this.drainPendingTurn();
      }
    }
  }

  /**
   * The PLANNER pre-stage (EXPERIENCE.md §F): for a COMPLEX free-text actuating request, stream a
   * planner turn, parse a {@link CommandPlan}, and stage it for the user's one-tap confirm BEFORE the
   * executor runs. A plan with `clarify` lines is surfaced as a question (no execution). A planner
   * that yields no parseable plan degrades to running the executor directly — the executor has its
   * own fail-closed gate, so the user's command still works. The planner turn neither reads nor
   * writes the document.
   */
  async proposePlan(
    task: string,
    grounding?: ResolvedGrounding,
    displayText?: string,
  ): Promise<void> {
    const t = task.trim();
    if (!t) return;
    if (this.state.busy) {
      // Mid-stream: don't lose it — queue as a commands turn (the executor, still gated).
      this.turnQueue.enqueue({ mode: 'commands', task: t, ...(grounding ? { grounding } : {}) });
      return;
    }
    const userMsg: ChatMessage = { id: this.id('u'), role: 'user', text: displayText ?? t };
    this.beginTurn({
      messages: [...this.state.messages, userMsg],
      pendingPlanClarification: undefined,
    });
    const controller = new AbortController();
    this.inflight = controller;
    try {
      const { plan, needsClarification } = await this.session.plan(t, {
        signal: controller.signal,
        ...(grounding ? { grounding } : {}),
      });
      if (controller.signal.aborted) return;
      if (!plan) {
        // No parseable plan → run the executor directly (it stages its own effect-level gate).
        this.set({ busy: false });
        void this.runCommands(t, grounding);
        return;
      }
      if (needsClarification) {
        this.set({
          messages: [
            ...this.state.messages,
            {
              id: this.id('a'),
              role: 'assistant',
              text: 'Before I plan this, I need one detail.',
            },
          ],
          pendingPlanClarification: {
            task: t,
            questions: plan.clarify,
            ...(grounding ? { grounding } : {}),
          },
          busy: false,
        });
        return;
      }
      if (deriveOutput(plan.intent) === 'chat') {
        this.set({
          messages: this.state.messages.filter((message) => message.id !== userMsg.id),
          busy: false,
        });
        void this.send(t, grounding);
        return;
      }
      this.set({
        pendingCommandPlan: { plan, task: t, ...(grounding ? { grounding } : {}) },
        busy: false,
      });
    } catch (err) {
      if (controller.signal.aborted || isAbortError(err)) this.set({ busy: false });
      else this.set({ error: errorText(err), busy: false });
    } finally {
      if (this.inflight === controller) this.inflight = undefined;
    }
  }

  /**
   * Resume the planner after it asked for clarification. The visible user message is the user's
   * answer, while the planner receives the original request plus that answer as structured context.
   */
  answerPlanClarification(answer: string): void {
    const a = answer.trim();
    if (!a) return;
    const pending = this.state.pendingPlanClarification;
    if (!pending) {
      void this.send(a);
      return;
    }
    const task = `${pending.task}\n\nUser clarification:\n${a}`;
    this.set({ pendingPlanClarification: undefined });
    void this.proposePlan(task, pending.grounding, a);
  }

  /**
   * Confirm the staged planner {@link CommandPlan}: clear it and run the executor on the normalized,
   * user-confirmed plan — which then stages its OWN effect-level gate (`pendingPlan`) before anything
   * actuates.
   */
  confirmCommandPlan(): void {
    const p = this.state.pendingCommandPlan;
    if (!p) return;
    this.set({ pendingCommandPlan: undefined });
    void this.runCommands(
      renderConfirmedPlanTask(p),
      p.grounding,
      renderConfirmedPlanDisplayText(p),
    );
  }

  /** Discard the staged planner plan without running anything (fail-closed: nothing executes). */
  cancelCommandPlan(): void {
    if (!this.state.pendingCommandPlan) return;
    this.set({ pendingCommandPlan: undefined });
  }

  /**
   * Cancel the in-flight turn's network/stream. No-op when idle. A queued turn (the `turnQueue` slot) still
   * drains afterwards THROUGH ITS OWN ROUTE — cancelling the current turn should run the one the user
   * lined up behind it, in the mode they requested, not discard it or downgrade it. Office.js host
   * writes already under way are not aborted (not abortable); this targets the assist stream only.
   */
  cancel(): void {
    this.inflight?.abort();
    // A loop gated on the user when cancelled must release fail-closed, so its awaited approval
    // resolves `false` and nothing actuates after the abort — for both the per-write gate and the
    // ADR-0005 plan-level gate.
    this.approvals.releaseAwaiting();
  }

  // ---- command loop (ADR-0004) --------------------------------------------

  /**
   * Drive the ADR-0004 read-many/write-one command loop with UI-backed, per-write approval. Streams
   * the grounded answer (tokens/citations) into an assistant message exactly like `send()`, while
   * narrating the loop's mechanics as `steps`. EVERY model-emitted write is staged as a
   * `pendingWrite` and actuates only after the user calls `approvePendingWrite()` — the fail-closed
   * human-in-the-loop. `send()`/`ask()` stay intact and untouched.
   */
  async runCommands(
    task: string,
    grounding?: ResolvedGrounding,
    displayText?: string,
  ): Promise<void> {
    const t = task.trim();
    if (!t) return;
    // Queued AS A COMMANDS turn (Finding #3): it drains back through runCommands — the fail-closed
    // plan/approval loop — NEVER through send(). A queued write turn is never downgraded to chat.
    if (this.state.busy) {
      this.turnQueue.enqueue({
        mode: 'commands',
        task: t,
        ...(grounding ? { grounding } : {}),
        ...(displayText ? { displayText } : {}),
      });
      return;
    }

    const userMsg: ChatMessage = { id: this.id('u'), role: 'user', text: displayText ?? t };
    const reply: ChatMessage = { id: this.id('a'), role: 'assistant', text: '', streaming: true };
    this.beginTurn({ messages: [...this.state.messages, userMsg, reply], steps: [] });

    const controller = new AbortController();
    this.inflight = controller;
    const sources: SourceRef[] = [];

    // The per-write approver (ADR-0004) and the plan-level approver (ADR-0005) both delegate to the
    // ApprovalCoordinator, which owns the staged-decision state and is fail-closed by construction.
    const approveWrite: ApproveWrite = (request) =>
      this.approvals.awaitWrite(
        { changeId: request.changeId, kind: request.kind, command: renderCommandLine(request) },
        request.changeId,
      );

    const approvePlan: ApprovePlan = (effects) =>
      this.approvals.awaitPlan({ effects, summary: summarizeEffects(effects) });

    const opts: RunCommandsOptions = {
      signal: controller.signal,
      approveWrite,
      approvePlan,
      ...(grounding ? { grounding } : {}),
    };

    try {
      for await (const ev of this.session.runCommands(t, opts)) {
        this.reduceLoopEvent(ev, reply.id, sources);
      }
    } catch (err) {
      if (controller.signal.aborted || isAbortError(err)) {
        this.patchMessage(reply.id, () => ({ cancelled: true }));
        this.currentTurnProvenance = undefined;
      } else {
        this.patchMessage(reply.id, () => ({ error: errorText(err) }));
        this.addStep('error', errorText(err));
      }
    } finally {
      // Finding #6: release any awaiting decision fail-closed AND drop both cards on EVERY terminal
      // path — including ones where the decision was ALREADY consumed (an approval whose execution
      // then THREW, or a write with no write-result), so no card lingers after the loop returns/throws.
      this.approvals.releaseAll();
      if (this.inflight === controller) this.inflight = undefined;
      this.patchMessage(reply.id, () => ({ streaming: false }));
      this.set({ busy: false, changes: this.store.list() });
      this.drainPendingTurn();
    }
  }

  /**
   * Run an explicit pasted CLI program (`set …`, `chart …`, `spill …`) through the same fail-closed
   * plan/approval loop as model-emitted commands, but without a Gemini echo turn. This is the manual
   * command escape hatch for power users and for copied examples from docs.
   */
  async runDirectCommands(program: string): Promise<void> {
    const p = program.trim();
    if (!p) return;
    if (this.state.busy) {
      this.turnQueue.enqueue({ mode: 'direct-commands', program: p });
      return;
    }

    const userMsg: ChatMessage = { id: this.id('u'), role: 'user', text: p };
    const reply: ChatMessage = { id: this.id('a'), role: 'assistant', text: '', streaming: true };
    this.beginTurn({ messages: [...this.state.messages, userMsg, reply], steps: [] });

    const controller = new AbortController();
    this.inflight = controller;
    const sources: SourceRef[] = [];

    const approveWrite: ApproveWrite = (request) =>
      this.approvals.awaitWrite(
        { changeId: request.changeId, kind: request.kind, command: renderCommandLine(request) },
        request.changeId,
      );
    const approvePlan: ApprovePlan = (effects) =>
      this.approvals.awaitPlan({ effects, summary: summarizeEffects(effects) });
    const opts: RunCommandsOptions = {
      signal: controller.signal,
      approveWrite,
      approvePlan,
      maxCommandsPerTurn: DIRECT_COMMAND_LIMIT,
      maxWritesPerTurn: DIRECT_COMMAND_LIMIT,
    };

    try {
      const runner = this.session.runCommandProgram
        ? this.session.runCommandProgram(p, opts)
        : this.session.runCommands(
            `Run this exact command program:\n\`\`\`cmd\n${p}\n\`\`\``,
            opts,
          );
      for await (const ev of runner) {
        this.reduceLoopEvent(ev, reply.id, sources);
      }
    } catch (err) {
      if (controller.signal.aborted || isAbortError(err)) {
        this.patchMessage(reply.id, () => ({ cancelled: true }));
        this.currentTurnProvenance = undefined;
      } else {
        this.patchMessage(reply.id, () => ({ error: errorText(err) }));
        this.addStep('error', errorText(err));
      }
    } finally {
      this.approvals.releaseAll();
      if (this.inflight === controller) this.inflight = undefined;
      this.patchMessage(reply.id, () => ({ streaming: false }));
      this.set({ busy: false, changes: this.store.list() });
      this.drainPendingTurn();
    }
  }

  /**
   * Reduce ONE command-loop event (the merged `SseEvent | CommandLoopEvent` union) into panel state.
   * Finding #6: handled EXHAUSTIVELY with an {@link assertNever} terminator, so a new event variant
   * is a compile error here rather than a silently-dropped step. SSE variants the run transcript does
   * not surface (`finding`/`slide`/`grounding-support`/`policy`/`related-questions`) are explicitly
   * ignored, not swallowed by a permissive `default`.
   */
  private reduceLoopEvent(
    ev: SseEvent | CommandLoopEvent,
    replyId: string,
    sources: SourceRef[],
  ): void {
    switch (ev.type) {
      case 'token':
        this.patchMessage(replyId, (m) => ({ text: m.text + ev.text }));
        return;
      case 'citation':
        sources.push(ev.source);
        this.patchMessage(replyId, () => ({ sources: [...sources] }));
        return;
      case 'provenance':
        // Finding #4: turn-scoped capture, threaded to writes by the runtime; cleared at turn end.
        this.currentTurnProvenance = ev.payload;
        return;
      case 'error':
        // `SseEvent` error (stream-level): the CommandLoopEvent union has no `error` variant, so
        // this narrows to the SSE shape with `code`/`message`.
        this.patchMessage(replyId, () => ({ error: ev.message }));
        this.addStep('error', ev.message);
        return;
      case 'code-execution':
        this.addStep('code-execution', 'Python code execution requested');
        return;
      case 'code-execution-result':
        this.addStep('code-execution', codeExecutionResultText(ev));
        return;
      case 'activity':
        this.addStep('activity', ev.text);
        return;
      case 'turn-start':
        this.addStep('turn-start', `Turn ${ev.turn}`);
        return;
      case 'command':
        this.addStep('command', commandStepText(ev));
        return;
      case 'expr-result':
        this.addStep('command', exprStepText(ev));
        return;
      case 'skill-registered':
        this.addStep('command', ev.result.message);
        return;
      case 'skill-expanded':
        this.addStep('command', `${ev.name} → ${ev.lines.length} line(s)`);
        return;
      case 'read-result':
        {
          const artifact = workspaceStepArtifact(ev.result);
          this.addStep(
            'read-result',
            artifact ? workspaceStepText(ev.result) : `read ${ev.intentLabel}`,
            artifact,
          );
        }
        return;
      case 'plan-preview':
        this.addStep('plan-preview', summarizeEffects(ev.effects));
        return;
      case 'write-result':
        this.addStep('write-result', writeStepText(ev));
        // The decision has been consumed by the loop; drop the staged pending-write card.
        this.approvals.consumeWriteResult();
        return;
      case 'no-fence':
        this.addStep('no-fence', `Turn ${ev.turn}: no command block — re-prompting`);
        return;
      case 'capped':
        this.addStep('capped', ev.reason);
        return;
      case 'done':
        this.addStep('done', 'Done');
        return;
      case 'exhausted':
        this.addStep('exhausted', `Stopped after ${ev.turns} turns`);
        return;
      // SSE variants the command-loop transcript does not surface — explicitly ignored.
      case 'finding':
      case 'slide':
      case 'grounding-support':
      case 'policy':
      case 'related-questions':
        return;
      default:
        assertNever(ev);
    }
  }

  /**
   * Approve the staged write — resolves the loop's `approveWrite` with `true` so it actuates.
   * Finding #6: pass the `changeId` the card was showing; a decision whose id no longer matches the
   * currently-staged write (a late click on a SUPERSEDED card) is IGNORED, so it can never apply a
   * request the loop has already moved past. Omit the id to approve whatever is staged (legacy).
   */
  approvePendingWrite(changeId?: ChangeId): void {
    this.approvals.approveWrite(changeId);
  }

  /** Reject the staged write — resolves the loop's `approveWrite` with `false`; nothing actuates. */
  rejectPendingWrite(changeId?: ChangeId): void {
    this.approvals.rejectWrite(changeId);
  }

  /**
   * Approve the staged composed plan (ADR-0005) — resolves the loop's `approvePlan` with `true`, so
   * the executor runs the previewed effect-set (each effect still gated/provenanced one-by-one).
   */
  approvePlan(): void {
    this.approvals.approvePlan();
  }

  /** Reject the staged plan — resolves `approvePlan` with `false`; the WHOLE plan is blocked. */
  rejectPlan(): void {
    this.approvals.rejectPlan();
  }

  private addStep(kind: RunStep['kind'], text: string, artifact?: RunStepArtifact): void {
    this.set({
      steps: [
        ...this.state.steps,
        { id: this.id('step'), kind, text, ...(artifact ? { artifact } : {}) },
      ],
    });
  }

  // ---- actuation review ---------------------------------------------------

  /**
   * Stage a reviewable, reversible change for the user to confirm. Finding #4: the proposal CAPTURES
   * the current turn's provenance AT CREATION and carries it explicitly — so when it is applied later
   * it is attributed to the turn that actually produced it, never to whatever turn happens to be
   * current at apply-time. A proposal created by a provenance-less turn carries none.
   */
  propose(kind: ActuationRequest['kind'], params: ActuationParams, label: string): Proposal {
    const proposal: Proposal = {
      changeId: asChangeId(this.id('c')),
      kind,
      params,
      label,
      status: 'pending',
      ...(this.currentTurnProvenance ? { provenance: this.currentTurnProvenance } : {}),
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

    // Finding #4: stamp the PROPOSAL's OWN captured provenance — never an ambient `lastProvenance`
    // a later turn could have overwritten. The session also receives it explicitly so the durable
    // host-metadata record is attributed to the turn that produced this change.
    const result = await this.session.apply(
      proposal.kind,
      proposal.params,
      proposal.changeId,
      proposal.provenance,
    );
    this.store.record(result, proposal.provenance);

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

  async refreshConversations(): Promise<void> {
    if (!this.session.listConversations) {
      this.set({
        conversations: {
          items: [],
          loading: false,
          loaded: true,
          error: 'Conversation history is unavailable in this build.',
        },
      });
      return;
    }
    const loadingState = { ...this.state.conversations };
    delete loadingState.error;
    this.set({
      conversations: {
        ...loadingState,
        loading: true,
      },
    });
    try {
      const result = await this.session.listConversations({ pageSize: 20 });
      this.set({
        conversations: {
          items: result.conversations.map((item) => this.toConversationItem(item)),
          loading: false,
          loaded: true,
          ...(result.nextPageToken ? { nextPageToken: result.nextPageToken } : {}),
        },
      });
    } catch (err) {
      this.set({
        conversations: {
          ...this.state.conversations,
          loading: false,
          loaded: true,
          error: errorText(err),
        },
      });
    }
  }

  resumeConversation(name: string): void {
    if (!this.session.resumeSession) return;
    this.session.resumeSession(name);
    this.set({
      conversations: {
        ...this.state.conversations,
        items: this.state.conversations.items.map((item) => ({
          ...item,
          active: item.name === name || item.id === name,
        })),
      },
    });
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

  /**
   * Record the skills (agents) and federated data stores discovered at boot via
   * `:listAvailableAgentViews` + the engine's data-store config (Task 6). Purely a view presenter —
   * it stores what was discovered so a future `@`-picker UI can read it from `getState()`; it adds
   * no execution or gate logic of its own.
   */
  setDiscoveredCatalog(agents: AgentView[], dataStores: EngineDataStore[]): void {
    this.set({ availableAgents: agents, availableDataStores: dataStores });
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
    // Finding #3: if a turn is in flight, queue AS A SKILL so the drain re-invokes invokeSkill (which
    // re-renders the call against the live registry) rather than collapsing it into a stale string.
    if (this.state.busy) {
      this.turnQueue.enqueue({ mode: 'skill', name, args });
      return;
    }
    await this.runCommands(renderSkillCall(skill, args));
  }

  // ---- internals ----------------------------------------------------------

  /**
   * Open a turn: mark busy, clear the prior turn's error/suggestions, and RESET the turn-local
   * provenance (Finding #4) so no leftover from an earlier turn can leak into this one's proposals.
   * Callers pass the turn's initial message/step patch.
   */
  private beginTurn(patch: Partial<PanelState>): void {
    this.currentTurnProvenance = undefined;
    this.set({ busy: true, error: undefined, suggestions: [], ...patch });
  }

  /**
   * Drain the single queued turn (Finding #3), if any, THROUGH ITS OWN ROUTE — ask→send,
   * commands→runCommands, skill→invokeSkill — so a queued mode is preserved end-to-end. Clearing the
   * slot first avoids re-enqueueing it; each dispatch is scheduled (not awaited) so draining does not
   * re-enter synchronously while the just-settled turn is unwinding. Exhaustive via {@link assertNever}.
   */
  private drainPendingTurn(): void {
    this.turnQueue.drain({
      ask: (query, grounding) => void this.send(query, grounding),
      commands: (task, grounding, displayText) =>
        void this.runCommands(task, grounding, displayText),
      directCommands: (program) => void this.runDirectCommands(program),
      skill: (name, args) => void this.invokeSkill(name, args),
    });
  }

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

  private toConversationItem(item: ConversationSummary): ConversationItem {
    const current = this.session.sessionId;
    const active = current === item.name || current === item.id;
    return {
      name: item.name,
      id: item.id,
      title: item.title,
      turnCount: item.turnCount,
      isPinned: item.isPinned,
      active,
      ...(item.state ? { state: item.state } : {}),
      ...(item.startedAt ? { startedAt: item.startedAt } : {}),
      ...(item.endedAt ? { endedAt: item.endedAt } : {}),
      ...(item.updatedAt ? { updatedAt: item.updatedAt } : {}),
    };
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

function workspaceStepText(result: unknown): string {
  if (!isWorkspaceResult(result)) return 'workspace';
  switch (result.workspace) {
    case 'list':
      return `workspace · ${result.artifacts.length} artifact${
        result.artifacts.length === 1 ? '' : 's'
      }`;
    case 'summary':
      return `workspace · ${result.artifact.name}`;
    case 'save':
      return `saved ${result.artifact.name} · ${formatBytes(result.artifact.bytes)}`;
    case 'cat':
      return `preview ${result.artifact.name}`;
    case 'grep':
      return `grep ${result.artifact.name} · ${result.matches.length} match${
        result.matches.length === 1 ? '' : 'es'
      }`;
    case 'cp':
      return `copied to ${result.artifact.name}`;
    case 'mv':
      return `renamed to ${result.artifact.name}`;
    case 'rm':
      return `deleted ${result.name}`;
    case 'error':
      return `workspace error: ${result.error}`;
  }
}

function workspaceStepArtifact(result: unknown): RunStepArtifact | undefined {
  if (!isWorkspaceResult(result)) return undefined;
  switch (result.workspace) {
    case 'list':
      return {
        title: 'Workspace artifacts',
        meta: result.artifacts.map(
          (a) => `${a.id} · ${a.name} · ${a.kind} · ${formatBytes(a.bytes)}`,
        ),
      };
    case 'summary':
    case 'save':
    case 'cat':
      return {
        title: `${result.artifact.id} · ${result.artifact.name}`,
        meta: artifactMeta(result.artifact),
        preview: result.preview,
      };
    case 'grep':
      return {
        title: `${result.artifact.id} · ${result.artifact.name}`,
        meta: [...artifactMeta(result.artifact), `pattern: ${result.pattern}`],
        matches: result.matches,
      };
    case 'cp':
    case 'mv':
      return {
        title: `${result.artifact.id} · ${result.artifact.name}`,
        meta: artifactMeta(result.artifact),
      };
    case 'rm':
      return { title: `deleted ${result.name}`, meta: [] };
    case 'error':
      return { title: 'Workspace error', meta: [result.error] };
  }
}

function isWorkspaceResult(result: unknown): result is WorkspaceResult {
  return !!result && typeof result === 'object' && 'workspace' in result;
}

function artifactMeta(artifact: WorkspaceArtifactSummary): string[] {
  return [
    `${artifact.kind} · ${artifact.mimeType}`,
    `${artifact.lineCount} line${artifact.lineCount === 1 ? '' : 's'} · ${formatBytes(artifact.bytes)}`,
    `source: ${artifact.sourceLabel}`,
    ...(artifact.truncated ? ['truncated'] : []),
  ];
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/** A one-line label for a `command` loop step: the parsed verb (or the corrective parse error). */
function commandStepText(ev: Extract<CommandLoopEvent, { type: 'command' }>): string {
  if ('error' in ev.compiled) return `error: ${ev.compiled.error}`;
  return ev.command.verb;
}

/** A one-line label for an `expr-result` loop step: the evaluated value, or the corrective error. */
function exprStepText(ev: Extract<CommandLoopEvent, { type: 'expr-result' }>): string {
  return 'error' in ev.result ? `expr error: ${ev.result.error}` : 'expr evaluated';
}

/** A one-line label for a `write-result` loop step: the write kind + its outcome. */
function writeStepText(ev: Extract<CommandLoopEvent, { type: 'write-result' }>): string {
  const r = ev.result;
  const outcome = r.ok ? (r.degraded ? 'degraded' : 'applied') : (r.error?.code ?? 'failed');
  // Observability: the change landed but its provenance is not durably recorded — make it visible.
  // `provenanceMissing` (no payload at all → unattributed) is distinct from `provenanceDropped`
  // (had a record, failed to persist); both leave the write without a durable trace.
  const provenance = !r.ok
    ? ''
    : r.provenanceMissing
      ? ' (⚠ unattributed — no provenance)'
      : r.provenanceDropped
        ? ' (⚠ provenance not recorded)'
        : '';
  return `${r.kind} — ${outcome}${provenance}`;
}

/** A one-line label for Gemini Enterprise code execution telemetry. */
function codeExecutionResultText(ev: Extract<SseEvent, { type: 'code-execution-result' }>): string {
  switch (ev.outcome) {
    case 'OUTCOME_OK':
      return 'Python code execution completed';
    case 'OUTCOME_FAILED':
      return 'Python code execution failed';
    case 'OUTCOME_DEADLINE_EXCEEDED':
      return 'Python code execution timed out';
    case 'OUTCOME_UNSPECIFIED':
      return 'Python code execution returned an unspecified outcome';
  }
}

/** A fetch aborted via AbortSignal rejects with a DOMException/Error named 'AbortError'. */
function isAbortError(err: unknown): boolean {
  return err instanceof Error && err.name === 'AbortError';
}
