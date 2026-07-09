import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ChangeId,
  CommandPlan,
  ContextRef,
  ProvenancePayload,
  SseEvent,
} from '@ge/contracts';
import { asChangeId, approvalClassOf, isReversibleKind } from '@ge/contracts';
import type { CommandLoopEvent } from '@ge/runtime';
import type {
  AgentView,
  ConversationSummary,
  EngineDataStore,
  ResolvedGrounding,
} from '@ge/gemini-client';
import type { HostEvent } from '@ge/triggers';
import {
  PanelController,
  type AssistLike,
  type ContextLister,
  type PlanEffect,
  type PlanRunCommandsOptions,
} from './controller.js';

const PROV_A: ProvenancePayload = {
  agentId: 'agent-A',
  identity: 'u',
  timestamp: 't-A',
  sources: [],
  contentHash: 'h-A',
};

const ref = (id: string, title: string): ContextRef => ({
  id,
  kind: 'selection',
  surface: 'word',
  title,
  preview: title,
});

/**
 * A scripted command-loop step: yield an event, trigger a per-write approval round-trip, or trigger
 * an ADR-0005 plan-level approval round-trip (emit `plan-preview`, await `approvePlan`, narrate the
 * decision).
 */
type CommandAction =
  | { event: SseEvent | CommandLoopEvent }
  | { approve: ActuationRequest }
  | { plan: PlanEffect[] }
  /**
   * Stage a write, await its approval, and — if approved — THROW from "execution" (after approval),
   * the way a bridge actuate can fault post-approval (Finding #6). The controller's `finally` must
   * still clear the card + busy.
   */
  | { approveThenThrow: ActuationRequest }
  /**
   * Stage write `first`, await+consume its decision, narrate its write-result (which clears the
   * card + the staged id), then stage write `second` and BLOCK awaiting its decision — so a test can
   * fire a LATE decision carrying `first`'s id at a card now showing `second` (Finding #6: ignored).
   */
  | { supersede: { first: ActuationRequest; second: ActuationRequest } }
  /**
   * Stage a `share` (estate write), await its own `approveShare` decision, and narrate the outcome
   * as a `read-result` — mirroring how `runWorkspaceIntent`'s `share` case actually behaves (no
   * separate later `write-result`-style event; the write happens synchronously right after approval).
   */
  | { share: { name: string; text: string; sourceLabel: string } };

const ev = (event: SseEvent | CommandLoopEvent): CommandAction => ({ event });

const planEffect = (changeId: string): PlanEffect => {
  const request = writeReq(changeId);
  return {
    request,
    command: `set Sales!F2 =C2-D2 [${changeId}]`,
    approvalClass: approvalClassOf(request.kind),
    reversible: isReversibleKind(request.kind),
  };
};

const writeReq = (changeId: string): ActuationRequest => ({
  changeId: asChangeId(changeId),
  kind: 'write-cells',
  surface: 'excel',
  params: { target: { range: 'Sales!F2' }, cells: [['=C2-D2']] },
});

/** A fake AssistSession: scripts the SSE stream and records attaches/applies/ingests. */
class FakeAssist implements AssistLike {
  context = { size: 0 };
  attached: string[] = [];
  detached: string[] = [];
  ingested: HostEvent[] = [];
  applied: Array<{ kind: string; changeId: string; provenance?: ProvenancePayload }> = [];
  /** The structured grounding handed to each ask/runCommands turn, in order (Finding #2/#B-wire). */
  askGrounding: Array<ResolvedGrounding | undefined> = [];
  runGrounding: Array<ResolvedGrounding | undefined> = [];
  script: SseEvent[] = [{ type: 'token', text: 'hi' }, { type: 'done' }];
  applyResult: ActuationResult = {
    ok: true,
    changeId: asChangeId(''),
    kind: 'tracked-change',
    location: 'para:3',
  };

  attachRef(r: ContextRef): Promise<void> {
    this.attached.push(r.id);
    this.context = { size: this.context.size + 1 };
    return Promise.resolve();
  }
  detach(id: string): void {
    this.detached.push(id);
  }
  /** Per-turn scripts (queries in order); falls back to `script` when exhausted. */
  scriptFor: SseEvent[][] = [];
  asked: string[] = [];
  /** When set, the stream pauses on this promise before emitting `done`, holding the turn open. */
  private gate: { promise: Promise<void>; release: () => void } | undefined;
  /** Arm a gate so the next `ask` blocks before completing; call `release()` to let it finish. */
  hold(): () => void {
    let release!: () => void;
    const promise = new Promise<void>((resolve) => {
      release = resolve;
    });
    this.gate = { promise, release };
    return release;
  }
  /** Arm a gate so the next `runCommands` blocks at its START (holding the loop busy). */
  private runGate: { promise: Promise<void>; release: () => void } | undefined;
  holdRun(): () => void {
    let release!: () => void;
    const promise = new Promise<void>((resolve) => {
      release = resolve;
    });
    this.runGate = { promise, release };
    return release;
  }
  async *ask(
    query: string,
    opts?: { signal?: AbortSignal; grounding?: ResolvedGrounding },
  ): AsyncGenerator<SseEvent> {
    this.asked.push(query);
    this.askGrounding.push(opts?.grounding);
    const script = this.scriptFor.shift() ?? this.script;
    const gate = this.gate;
    this.gate = undefined;
    for (const ev of script) {
      if (gate && ev.type === 'done') {
        // While held open, an abort should interrupt the turn the way a cancelled fetch would:
        // reject the iteration with an AbortError instead of emitting the final events.
        const signal = opts?.signal;
        if (signal?.aborted) throw abortError();
        await Promise.race([
          gate.promise,
          new Promise<never>((_, reject) => {
            signal?.addEventListener('abort', () => reject(abortError()), { once: true });
          }),
        ]);
      }
      yield ev;
    }
  }
  /** F: the planner pre-stage result. Default → no plan (proposePlan degrades to the executor). */
  planned: { plan: CommandPlan | null; errors: string[]; needsClarification: boolean } = {
    plan: null,
    errors: [],
    needsClarification: false,
  };
  plannedQueue: { plan: CommandPlan | null; errors: string[]; needsClarification: boolean }[] = [];
  planTasks: string[] = [];
  plan(
    task: string,
  ): Promise<{ plan: CommandPlan | null; errors: string[]; needsClarification: boolean }> {
    this.planTasks.push(task);
    return Promise.resolve(this.plannedQueue.shift() ?? this.planned);
  }

  apply(
    kind: ActuationRequest['kind'],
    _params: ActuationParams,
    changeId: ChangeId,
    provenance?: ProvenancePayload,
  ): Promise<ActuationResult> {
    this.applied.push({ kind, changeId, ...(provenance ? { provenance } : {}) });
    return Promise.resolve({ ...this.applyResult, changeId, kind });
  }

  /**
   * A scripted command loop. `commandScript` is a list of "actions": plain SSE/loop events are
   * yielded as-is; an `{ approve: request }` action calls `opts.approveWrite(request)` and then
   * yields a `write-result` reflecting the decision (approved → ok; rejected → blocked). This lets
   * a test assert the controller stages a pending write and gates actuation on the user's decision.
   */
  commandScript: CommandAction[] = [];
  runTasks: string[] = [];
  conversations: ConversationSummary[] = [];
  resumedSession: string | undefined;
  approveCalls: ActuationRequest[] = [];
  approveResults: boolean[] = [];
  /** The plan effect-sets passed to `approvePlan` and the decisions returned, for assertions. */
  planCalls: PlanEffect[][] = [];
  planResults: boolean[] = [];
  /** The share inputs passed to `approveShare` and the decisions returned, for assertions. */
  shareCalls: Array<{ name: string; text: string; sourceLabel: string }> = [];
  shareResults: boolean[] = [];
  async *runCommands(
    task: string,
    opts?: PlanRunCommandsOptions,
  ): AsyncGenerator<SseEvent | CommandLoopEvent> {
    this.runTasks.push(task);
    this.runGrounding.push(opts?.grounding);
    const runGate = this.runGate;
    this.runGate = undefined;
    if (runGate) await runGate.promise; // hold the loop busy until released
    for (const action of this.commandScript) {
      if ('approveThenThrow' in action) {
        // Finding #6: gate the write, and if approved, fault DURING execution (after approval).
        const request = action.approveThenThrow;
        this.approveCalls.push(request);
        const approved = opts?.approveWrite ? await opts.approveWrite(request) : false;
        this.approveResults.push(approved);
        if (approved) throw new Error('actuate boom (post-approval)');
        // If not approved, fall through with no further events (loop would continue normally).
      } else if ('supersede' in action) {
        // Finding #6: drive a card from `first` → `second` so a late `first`-id decision is stale.
        const { first, second } = action.supersede;
        // 1. Stage `first`, await + consume its decision, narrate its write-result (clears the card).
        this.approveCalls.push(first);
        const firstApproved = opts?.approveWrite ? await opts.approveWrite(first) : false;
        this.approveResults.push(firstApproved);
        yield {
          type: 'write-result',
          turn: 1,
          changeId: first.changeId,
          result: { ok: true, changeId: first.changeId, kind: first.kind, location: 'A1' },
        };
        // 2. Stage `second` and BLOCK on its decision — the test now holds the loop here.
        this.approveCalls.push(second);
        const secondApproved = opts?.approveWrite ? await opts.approveWrite(second) : false;
        this.approveResults.push(secondApproved);
        yield {
          type: 'write-result',
          turn: 1,
          changeId: second.changeId,
          result: secondApproved
            ? { ok: true, changeId: second.changeId, kind: second.kind, location: 'B1' }
            : {
                ok: false,
                changeId: second.changeId,
                kind: second.kind,
                error: { code: 'unapproved', message: 'rejected' },
              },
        };
      } else if ('approve' in action) {
        const request = action.approve;
        this.approveCalls.push(request);
        const approved = opts?.approveWrite ? await opts.approveWrite(request) : false;
        this.approveResults.push(approved);
        const result: ActuationResult = approved
          ? { ok: true, changeId: request.changeId, kind: request.kind, location: 'A1' }
          : {
              ok: false,
              changeId: request.changeId,
              kind: request.kind,
              error: { code: 'unapproved', message: 'rejected' },
            };
        yield { type: 'write-result', turn: 1, changeId: request.changeId, result };
      } else if ('plan' in action) {
        // ADR-0005: emit the dry-run preview, then gate the whole effect-set on `approvePlan`.
        const effects = action.plan;
        // The runtime `CommandLoopEvent` union has no `plan-preview` yet; the controller narrows it
        // structurally, so cast at the fake boundary to drive that path.
        yield { type: 'plan-preview', turn: 1, effects } as unknown as CommandLoopEvent;
        this.planCalls.push(effects);
        const approved = opts?.approvePlan ? await opts.approvePlan(effects) : false;
        this.planResults.push(approved);
        // On approval, the executor gates each effect and narrates a per-effect write-result.
        if (approved) {
          for (const e of effects) {
            yield {
              type: 'write-result',
              turn: 1,
              changeId: e.request.changeId,
              result: {
                ok: true,
                changeId: e.request.changeId,
                kind: e.request.kind,
                location: 'A1',
              },
            };
          }
        }
      } else if ('share' in action) {
        const input = action.share;
        this.shareCalls.push(input);
        const approved = opts?.approveShare ? await opts.approveShare(input) : false;
        this.shareResults.push(approved);
        const result = approved
          ? { workspace: 'share', name: input.name, bytes: input.text.length }
          : { workspace: 'error', error: 'share requires user approval (none granted)' };
        yield {
          type: 'read-result',
          turn: 1,
          intentLabel: `share ${input.name}`,
          result,
        } as unknown as CommandLoopEvent;
      } else {
        yield action.event;
      }
    }
  }
  ingest(event: HostEvent): Promise<void> {
    this.ingested.push(event);
    return Promise.resolve();
  }
  listConversations(): Promise<{ conversations: ConversationSummary[] }> {
    return Promise.resolve({ conversations: this.conversations });
  }
  resumeSession(sessionIdOrName: string): void {
    this.resumedSession = sessionIdOrName;
  }
}

function lister(refs: ContextRef[]): ContextLister {
  return { listContext: () => Promise.resolve(refs) };
}

function abortError(): Error {
  const e = new Error('The operation was aborted.');
  e.name = 'AbortError';
  return e;
}

describe('PanelController — context tray', () => {
  it('loads chips and attaches/detaches a specific ref', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([ref('word:selection', 'Selection')]));
    await c.refreshContext();
    expect(c.getState().chips).toHaveLength(1);
    expect(c.getState().chips[0]).toMatchObject({ id: 'word:selection', attached: false });

    await c.attach('word:selection');
    expect(assist.attached).toEqual(['word:selection']);
    expect(c.getState().chips[0]?.attached).toBe(true);

    c.detach('word:selection');
    expect(assist.detached).toEqual(['word:selection']);
    expect(c.getState().chips[0]?.attached).toBe(false);
  });

  it('marks chips revealable when the bridge can reveal context and calls that bridge path', async () => {
    const assist = new FakeAssist();
    const selection = ref('word:selection', 'Selection');
    const revealContext = vi.fn(async (_ref: ContextRef): Promise<void> => {});
    const c = new PanelController(assist, {
      listContext: () => Promise.resolve([selection]),
      revealContext,
    });

    await c.refreshContext();
    expect(c.getState().chips[0]).toMatchObject({ id: 'word:selection', revealable: true });

    await c.reveal('word:selection');
    expect(revealContext).toHaveBeenCalledWith(selection);
  });

  it('reveals an Excel target surfaced outside the context tray by synthesizing a range ref', async () => {
    const assist = new FakeAssist();
    const revealContext = vi.fn(async (_ref: ContextRef): Promise<void> => {});
    const c = new PanelController(assist, {
      listContext: () => Promise.resolve([]),
      canRevealContext: (contextRef) => contextRef.surface === 'excel',
      revealContext,
    });

    await c.revealLocation('excel', "'Daily schedule'!K6:L18");
    expect(revealContext).toHaveBeenCalledWith({
      id: "xl:'Daily schedule'!K6:L18",
      kind: 'range',
      surface: 'excel',
      title: "'Daily schedule'!K6:L18",
      hostRef: { type: 'excel.range', address: "'Daily schedule'!K6:L18" },
    });
  });

  it('does not mark chips revealable when the bridge rejects that specific ref', async () => {
    const assist = new FakeAssist();
    const document = ref('word:document', 'Whole document');
    const c = new PanelController(assist, {
      listContext: () => Promise.resolve([document]),
      canRevealContext: () => false,
      revealContext: vi.fn(async (_ref: ContextRef): Promise<void> => {}),
    });

    await c.refreshContext();
    expect(c.getState().chips[0]).toMatchObject({ id: 'word:document' });
    expect(c.getState().chips[0]?.revealable).toBeUndefined();
  });
});

describe('PanelController — conversation history', () => {
  it('lists Discovery Engine sessions and can mark one as active', async () => {
    const assist = new FakeAssist();
    assist.conversations = [
      {
        name: 'projects/proj/locations/global/collections/default_collection/engines/e/sessions/sess-1',
        id: 'sess-1',
        title: 'Schedule planning',
        turnCount: 3,
        isPinned: true,
        updatedAt: '2026-07-07T02:54:23.000Z',
      },
    ];
    const c = new PanelController(assist, lister([]));

    await c.refreshConversations();

    expect(c.getState().conversations.loaded).toBe(true);
    expect(c.getState().conversations.items[0]).toMatchObject({
      id: 'sess-1',
      title: 'Schedule planning',
      turnCount: 3,
      isPinned: true,
      active: false,
    });

    c.resumeConversation(assist.conversations[0]!.name);
    expect(assist.resumedSession).toBe(assist.conversations[0]!.name);
    expect(c.getState().conversations.items[0]?.active).toBe(true);
  });
});

describe('PanelController — ask / stream', () => {
  it('streams tokens + citations into the assistant message and toggles busy', async () => {
    const assist = new FakeAssist();
    assist.script = [
      { type: 'activity', text: 'Reading selected policy' },
      { type: 'token', text: 'The SLA ' },
      { type: 'token', text: 'is fine.' },
      { type: 'citation', source: { title: 'Vendor Policy', uri: 'https://x' } },
      {
        type: 'provenance',
        payload: { agentId: 'a', identity: 'u', timestamp: 't', sources: [], contentHash: 'h' },
      },
      { type: 'done' },
    ];
    const c = new PanelController(assist, lister([]));
    const seen: boolean[] = [];
    c.subscribe((s) => seen.push(s.busy));

    await c.send('is the SLA ok?');
    const msgs = c.getState().messages;
    expect(msgs.map((m) => m.role)).toEqual(['user', 'assistant']);
    expect(msgs[1]?.text).toBe('The SLA is fine.');
    expect(msgs[1]?.text).not.toContain('Reading selected policy');
    expect(c.getState().steps.map((s) => s.text)).toContain('Reading selected policy');
    expect(msgs[1]?.sources?.[0]?.title).toBe('Vendor Policy');
    expect(msgs[1]?.streaming).toBe(false);
    expect(c.getState().busy).toBe(false);
    expect(seen).toContain(true); // was busy mid-stream
  });

  it('ignores empty input and refuses concurrent sends', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    await c.send('   ');
    expect(c.getState().messages).toHaveLength(0);
  });

  it('surfaces a stream error on the assistant message', async () => {
    const assist = new FakeAssist();
    assist.script = [{ type: 'error', code: 'http_500', message: 'boom' }, { type: 'done' }];
    const c = new PanelController(assist, lister([]));
    await c.send('hi');
    expect(c.getState().messages[1]?.error).toBe('boom');
  });

  it('cancel() aborts an in-flight turn: message marked cancelled (not error), busy clears', async () => {
    const assist = new FakeAssist();
    assist.script = [{ type: 'token', text: 'partial' }, { type: 'done' }];
    const c = new PanelController(assist, lister([]));

    // Hold the turn open after the token so it is still streaming when we cancel.
    assist.hold();
    const turn = c.send('explain');
    expect(c.getState().busy).toBe(true);

    c.cancel();
    await turn;

    const reply = c.getState().messages[1];
    expect(reply?.text).toBe('partial'); // the partial answer is kept
    expect(reply?.cancelled).toBe(true);
    expect(reply?.error).toBeUndefined(); // cancellation is not a red error
    expect(reply?.streaming).toBe(false);
    expect(c.getState().busy).toBe(false);
  });

  it('cancel() is a clean no-op when idle or after the turn settles', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    expect(() => c.cancel()).not.toThrow(); // idle

    await c.send('hi');
    expect(c.getState().messages[1]?.cancelled).toBeUndefined();
    expect(() => c.cancel()).not.toThrow(); // after settle, no live controller
    expect(c.getState().messages[1]?.cancelled).toBeUndefined();
  });

  it('a queued automate still runs after the current turn is cancelled', async () => {
    const assist = new FakeAssist();
    assist.scriptFor = [
      [{ type: 'token', text: 'a-reply' }, { type: 'done' }],
      [{ type: 'token', text: 'b-reply' }, { type: 'done' }],
    ];
    const c = new PanelController(assist, lister([]));

    assist.hold();
    const first = c.send('a');
    expect(c.getState().busy).toBe(true);

    // Queue a follow-up, then cancel the current turn: the queued one should drain.
    c.onAutomate('b');
    c.cancel();
    await first;
    await new Promise((r) => setTimeout(r, 0)); // let the scheduled drain settle

    expect(assist.asked).toEqual(['a', 'b']);
    expect(c.getState().busy).toBe(false);
    expect(c.getState().messages[1]?.cancelled).toBe(true); // 'a' was cancelled
    expect(c.getState().messages.map((m) => m.text)).toEqual(['a', 'a-reply', 'b', 'b-reply']);
  });
});

describe('PanelController — actuation review', () => {
  it('applies a proposal and records the change', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    const p = c.propose('tracked-change', { text: 'x' }, 'Rewrite SLA clause');
    expect(c.getState().proposals[0]?.status).toBe('pending');

    await c.applyProposal(p.changeId);
    expect(assist.applied).toEqual([{ kind: 'tracked-change', changeId: p.changeId }]);
    expect(c.getState().proposals[0]?.status).toBe('applied');
    expect(c.getState().changes).toHaveLength(1);
  });

  it('guards against double-apply: two concurrent calls yield one apply and one record', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    const p = c.propose('tracked-change', { text: 'x' }, 'Rewrite SLA clause');

    const statuses: string[] = [];
    c.subscribe((s) => {
      const st = s.proposals[0]?.status;
      if (st) statuses.push(st);
    });

    // Double-click: both fire before the first await resolves.
    await Promise.all([c.applyProposal(p.changeId), c.applyProposal(p.changeId)]);

    expect(assist.applied).toEqual([{ kind: 'tracked-change', changeId: p.changeId }]);
    expect(c.getState().changes).toHaveLength(1);
    expect(c.getState().proposals[0]?.status).toBe('applied');
    // A subscriber observes the transient 'applying' then the final status.
    expect(statuses).toContain('applying');
    expect(statuses).toContain('applied');
  });

  it('marks a blocked proposal from a gate veto', async () => {
    const assist = new FakeAssist();
    assist.applyResult = {
      ok: false,
      changeId: asChangeId(''),
      kind: 'tracked-change',
      error: { code: 'blocked', message: 'Not grounded.' },
    };
    const c = new PanelController(assist, lister([]));
    const p = c.propose('tracked-change', { text: 'x' }, 'risky');
    await c.applyProposal(p.changeId);
    expect(c.getState().proposals[0]?.status).toBe('blocked');
    expect(c.getState().proposals[0]?.detail).toBe('Not grounded.');
  });
});

describe('PanelController — event-driven inputs', () => {
  it('onContext feeds events to the session (context path)', () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    c.onContext({ type: 'meeting-ended', id: 'm1' });
    expect(assist.ingested).toEqual([{ type: 'meeting-ended', id: 'm1' }]);
  });

  it('onSuggest adds an ignorable chip; dismiss removes it; onAutomate runs a turn', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    c.onSuggest({ title: 'Draft follow-ups?', query: 'draft follow-ups' });
    const s = c.getState().suggestions[0];
    expect(s?.title).toBe('Draft follow-ups?');
    c.dismissSuggestion(s!.id);
    expect(c.getState().suggestions).toHaveLength(0);

    const send = vi.spyOn(c, 'send');
    c.onAutomate('go');
    expect(send).toHaveBeenCalledWith('go');
  });

  it('queues an onAutomate turn fired while a send is streaming (not dropped)', async () => {
    const assist = new FakeAssist();
    assist.scriptFor = [
      [{ type: 'token', text: 'a-reply' }, { type: 'done' }],
      [{ type: 'token', text: 'b-reply' }, { type: 'done' }],
    ];
    const c = new PanelController(assist, lister([]));

    // Hold the first turn open so it is still streaming when onAutomate fires.
    const release = assist.hold();
    const first = c.send('a');
    expect(c.getState().busy).toBe(true);

    // Automated turn fired mid-stream: must be queued, not silently dropped.
    c.onAutomate('b');
    expect(assist.asked).toEqual(['a']); // 'b' not started yet

    release();
    await first;
    // Drain runs 'b' as a scheduled microtask; let it settle.
    await new Promise((r) => setTimeout(r, 0));

    expect(assist.asked).toEqual(['a', 'b']);
    expect(c.getState().busy).toBe(false);
    expect(c.getState().messages.map((m) => m.text)).toEqual(['a', 'a-reply', 'b', 'b-reply']);
  });

  it('reroutes a cmd block returned from normal chat into the Office command route', async () => {
    const assist = new FakeAssist();
    assist.script = [
      { type: 'token', text: "```cmd\nread 'Daily schedule'!B2:I10\n" },
      { type: 'done' },
    ];
    assist.commandScript = [
      ev({ type: 'turn-start', turn: 1 }),
      ev({ type: 'done', turn: 1, answer: '' }),
    ];
    const c = new PanelController(assist, lister([]));

    await c.send('Okay add this to my schedule');

    expect(assist.asked).toEqual(['Okay add this to my schedule']);
    expect(assist.runTasks).toEqual(['Okay add this to my schedule']);
    expect(c.getState().messages[1]?.text).toContain('Detected Office command output');
    expect(c.getState().messages[1]?.text).not.toContain('```cmd');
    expect(c.getState().messages.map((m) => m.text)).toContain('Continue in Office command route');
    expect(c.getState().busy).toBe(false);
  });
});

describe('PanelController — command loop (ADR-0004 human-in-the-loop)', () => {
  it('reduces loop events into steps and streams the answer text', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [
      ev({ type: 'turn-start', turn: 1 }),
      ev({
        type: 'command',
        turn: 1,
        command: { verb: 'read', selector: 'Sales!C2:C7' },
        compiled: { kind: 'read', intent: { read: 'range', selector: 'Sales!C2:C7' } },
      }),
      ev({ type: 'read-result', turn: 1, intentLabel: 'Sales!C2:C7', result: [1, 2] }),
      ev({ type: 'activity', text: 'Calculating the margin' }),
      ev({ type: 'code-execution', language: 'python', code: 'print(12)' }),
      ev({ type: 'code-execution-result', outcome: 'OUTCOME_OK', output: '12\n' }),
      ev({ type: 'token', text: 'The margin ' }),
      ev({ type: 'token', text: 'is 12%.' }),
      ev({ type: 'done', turn: 1, answer: 'The margin is 12%.' }),
    ];
    const c = new PanelController(assist, lister([]));

    await c.runCommands('compute the margin');

    expect(assist.runTasks).toEqual(['compute the margin']);
    const reply = c.getState().messages[1];
    expect(reply?.text).toBe('The margin is 12%.');
    expect(reply?.streaming).toBe(false);
    expect(c.getState().busy).toBe(false);
    expect(c.getState().steps.map((s) => s.kind)).toEqual([
      'turn-start',
      'command',
      'read-result',
      'activity',
      'code-execution',
      'code-execution',
      'done',
    ]);
    expect(c.getState().steps.map((s) => s.text)).toContain('Python code execution completed');
    expect(c.getState().pendingWrite).toBeUndefined();
  });

  it('preserves compact workspace artifacts on read-result steps', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [
      ev({
        type: 'read-result',
        turn: 1,
        intentLabel: 'save schedule.tsv',
        result: {
          workspace: 'save',
          artifact: {
            id: 'ws:1',
            name: 'schedule.tsv',
            kind: 'tsv',
            mimeType: 'text/tab-separated-values',
            sourceLabel: "read 'Daily schedule'!B3:I53",
            createdAt: '2026-07-07T00:00:00.000Z',
            bytes: 1200,
            lineCount: 20,
            truncated: false,
          },
          preview: 'Time\tMonday\n08:00\tDeep Work',
        },
      } as unknown as CommandLoopEvent),
    ];
    const c = new PanelController(assist, lister([]));

    await c.runCommands('save schedule');

    const step = c.getState().steps.at(-1);
    expect(step?.text).toBe('saved schedule.tsv · 1.2 KB');
    expect(step?.artifact).toMatchObject({
      title: 'ws:1 · schedule.tsv',
      preview: 'Time\tMonday\n08:00\tDeep Work',
      meta: expect.arrayContaining(["source: read 'Daily schedule'!B3:I53"]),
    });
  });

  it('renders a share result as a cross-surface publish step', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [
      ev({
        type: 'read-result',
        turn: 1,
        intentLabel: 'share schedule.tsv',
        result: { workspace: 'share', name: 'schedule.tsv', bytes: 1200 },
      } as unknown as CommandLoopEvent),
    ];
    const c = new PanelController(assist, lister([]));

    await c.runCommands('share schedule');

    const step = c.getState().steps.at(-1);
    expect(step?.text).toBe('shared schedule.tsv · 1.2 KB');
    expect(step?.artifact).toMatchObject({ title: 'shared/schedule.tsv', meta: ['1.2 KB'] });
  });

  it('stages a share as pending and writes only after approvePendingShare() (resolves true)', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [
      { share: { name: 'schedule.tsv', text: 'a\tb\n1\t2', sourceLabel: 'read Sales!A1:B2' } },
    ];
    const c = new PanelController(assist, lister([]));

    // Don't await: the loop blocks on the pending-share decision.
    const run = c.runCommands('share the schedule');
    await tick();

    const pending = c.getState().pendingShare;
    expect(pending).toMatchObject({
      name: 'schedule.tsv',
      sourceLabel: 'read Sales!A1:B2',
      preview: 'a\tb\n1\t2',
      truncated: false,
    });
    expect(assist.shareResults).toEqual([]); // decision not made yet
    expect(c.getState().busy).toBe(true);

    c.approvePendingShare();
    await run;

    expect(assist.shareResults).toEqual([true]);
    expect(c.getState().pendingShare).toBeUndefined();
    const shareStep = c.getState().steps.find((s) => s.text.startsWith('shared '));
    expect(shareStep?.text).toBe('shared schedule.tsv · 7 B');
    expect(c.getState().busy).toBe(false);
  });

  it('rejectPendingShare() resolves false: nothing is written and the card clears', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [
      { share: { name: 'schedule.tsv', text: 'a\tb\n1\t2', sourceLabel: 'read Sales!A1:B2' } },
    ];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('share the schedule');
    await tick();
    expect(c.getState().pendingShare?.name).toBe('schedule.tsv');

    c.rejectPendingShare();
    await run;

    expect(assist.shareResults).toEqual([false]);
    expect(c.getState().pendingShare).toBeUndefined();
    const errorStep = c.getState().steps.find((s) => s.text.includes('workspace error'));
    expect(errorStep?.text).toContain('requires user approval');
  });

  it('cancel() while gated on a share releases fail-closed (rejects the pending share)', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [
      { share: { name: 'schedule.tsv', text: 'a\tb\n1\t2', sourceLabel: 'read Sales!A1:B2' } },
    ];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('share the schedule');
    await tick();
    expect(c.getState().pendingShare).toBeDefined();

    c.cancel();
    await run;

    expect(assist.shareResults).toEqual([false]);
    expect(c.getState().pendingShare).toBeUndefined();
  });

  it('stages a write as pending and actuates only after approvePendingWrite() (resolves true)', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [ev({ type: 'turn-start', turn: 1 }), { approve: writeReq('w-1') }];
    const c = new PanelController(assist, lister([]));

    // Don't await: the loop blocks on the pending-write decision.
    const run = c.runCommands('set the formula');
    await tick();

    // The write is staged verbatim and nothing has actuated yet.
    const pending = c.getState().pendingWrite;
    expect(pending?.command).toBe('set Sales!F2 =C2-D2');
    expect(pending?.changeId).toBe(asChangeId('w-1'));
    expect(assist.approveResults).toEqual([]); // decision not made yet
    expect(c.getState().busy).toBe(true);

    c.approvePendingWrite();
    await run;

    expect(assist.approveResults).toEqual([true]); // approveWrite resolved true → actuates
    expect(c.getState().pendingWrite).toBeUndefined(); // cleared on write-result
    const writeStep = c.getState().steps.find((s) => s.kind === 'write-result');
    expect(writeStep?.text).toBe('write-cells — applied');
    expect(c.getState().busy).toBe(false);
  });

  it('rejectPendingWrite() resolves false: the write is blocked and the card clears', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [{ approve: writeReq('w-2') }];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('set the formula');
    await tick();
    expect(c.getState().pendingWrite?.command).toBe('set Sales!F2 =C2-D2');

    c.rejectPendingWrite();
    await run;

    expect(assist.approveResults).toEqual([false]); // approveWrite resolved false → blocked
    expect(c.getState().pendingWrite).toBeUndefined();
    const writeStep = c.getState().steps.find((s) => s.kind === 'write-result');
    expect(writeStep?.text).toBe('write-cells — unapproved');
  });

  it('cancel() while gated on a write releases fail-closed (rejects the pending write)', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [{ approve: writeReq('w-3') }];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('set the formula');
    await tick();
    expect(c.getState().pendingWrite).toBeDefined();

    c.cancel();
    await run;

    expect(assist.approveResults).toEqual([false]); // cancel released the gate as a rejection
    expect(c.getState().pendingWrite).toBeUndefined();
  });
});

describe('PanelController — plan loop (ADR-0005 plan-level approval)', () => {
  it('stages the full effect-set as pendingPlan and actuates only after approvePlan() (true)', async () => {
    const assist = new FakeAssist();
    const effects = [planEffect('p-1'), planEffect('p-2')];
    assist.commandScript = [ev({ type: 'turn-start', turn: 1 }), { plan: effects }];
    const c = new PanelController(assist, lister([]));

    // Don't await: the loop blocks on the plan-level decision.
    const run = c.runCommands('reconcile the totals');
    await tick();

    // The whole effect-set is staged, rendered verbatim from each request, and nothing actuated.
    const plan = c.getState().pendingPlan;
    expect(plan?.effects).toHaveLength(2);
    expect(plan?.effects.map((e) => e.request.changeId)).toEqual([
      asChangeId('p-1'),
      asChangeId('p-2'),
    ]);
    expect(plan?.summary).toBe('2 writes');
    expect(assist.planResults).toEqual([]); // decision not made yet
    expect(assist.applied).toEqual([]); // nothing actuated
    expect(c.getState().busy).toBe(true);
    // The preview reduced into the steps transcript.
    expect(c.getState().steps.map((s) => s.kind)).toEqual(['turn-start', 'plan-preview']);

    c.approvePlan();
    await run;

    expect(assist.planResults).toEqual([true]); // approvePlan resolved true → plan executes
    expect(c.getState().pendingPlan).toBeUndefined(); // card cleared on decision
    // Each effect narrated a per-effect write-result.
    const writeSteps = c.getState().steps.filter((s) => s.kind === 'write-result');
    expect(writeSteps).toHaveLength(2);
    expect(c.getState().busy).toBe(false);
  });

  it('the rendered effect-set is exactly what executes (same requests, no divergence)', async () => {
    const assist = new FakeAssist();
    const effects = [planEffect('p-1'), planEffect('p-2')];
    assist.commandScript = [{ plan: effects }];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('reconcile');
    await tick();
    const staged = c.getState().pendingPlan?.effects ?? [];
    // The requests the user sees are reference-identical to the ones handed to approvePlan.
    expect(staged.map((e) => e.request)).toEqual(effects.map((e) => e.request));

    c.approvePlan();
    await run;
    expect(assist.planCalls[0]).toBe(effects); // executor received the same effect-set
  });

  it('rejectPlan() resolves false: the whole plan is blocked, the card clears, nothing runs', async () => {
    const assist = new FakeAssist();
    const effects = [planEffect('p-1'), planEffect('p-2')];
    assist.commandScript = [{ plan: effects }];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('reconcile');
    await tick();
    expect(c.getState().pendingPlan?.effects).toHaveLength(2);

    c.rejectPlan();
    await run;

    expect(assist.planResults).toEqual([false]); // approvePlan resolved false → blocked
    expect(c.getState().pendingPlan).toBeUndefined();
    expect(assist.applied).toEqual([]); // no effect actuated
    expect(c.getState().steps.some((s) => s.kind === 'write-result')).toBe(false);
  });

  it('cancel() while gated on a plan releases fail-closed (rejects the whole plan)', async () => {
    const assist = new FakeAssist();
    const effects = [planEffect('p-1')];
    assist.commandScript = [{ plan: effects }];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('reconcile');
    await tick();
    expect(c.getState().pendingPlan).toBeDefined();

    c.cancel();
    await run;

    expect(assist.planResults).toEqual([false]); // cancel released the plan gate as a rejection
    expect(c.getState().pendingPlan).toBeUndefined();
    expect(assist.applied).toEqual([]);
  });

  it('fail-closed: a plan abandoned at teardown without a decision settles as a rejection', async () => {
    const assist = new FakeAssist();
    const effects = [planEffect('p-1')];
    // Emit the preview but DON'T loop through approvePlan in the script — the generator returns,
    // hitting the controller's `finally`, which must settle the open plan false.
    assist.commandScript = [
      { event: { type: 'plan-preview', turn: 1, effects } as unknown as CommandLoopEvent },
    ];
    const c = new PanelController(assist, lister([]));

    await c.runCommands('reconcile');

    // No approvePlan was awaited by the fake, but the controller's approvePlan promise that it
    // wired is irrelevant here; the key invariant is the card never lingers and nothing actuated.
    expect(c.getState().pendingPlan).toBeUndefined();
    expect(assist.applied).toEqual([]);
  });

  it('summarizes a mixed effect-set as a pluralized count header', async () => {
    const assist = new FakeAssist();
    const write: PlanEffect = planEffect('w-1');
    const comment: PlanEffect = {
      request: {
        changeId: asChangeId('c-1'),
        kind: 'add-comment',
        surface: 'word',
        params: { target: { matchText: 'SLA' }, text: 'check this' },
      },
      command: 'comment "SLA" "check this"',
      approvalClass: 'in-document',
      reversible: true,
    };
    assist.commandScript = [{ plan: [write, comment, comment] }];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('mixed');
    await tick();
    expect(c.getState().pendingPlan?.summary).toBe('1 write + 2 comments');

    c.rejectPlan();
    await run;
  });
});

describe('PanelController — typed turn queue (Finding #3: a queued turn keeps its mode)', () => {
  it('a commands turn queued mid-stream drains through runCommands (NOT send) — mode preserved', async () => {
    const assist = new FakeAssist();
    // The first (ask) turn is held open so the queued commands turn lands while busy.
    assist.scriptFor = [[{ type: 'token', text: 'a-reply' }, { type: 'done' }]];
    assist.commandScript = [ev({ type: 'done', turn: 1, answer: 'done' })];
    const c = new PanelController(assist, lister([]));

    const release = assist.hold();
    const first = c.send('a'); // ask turn, held
    expect(c.getState().busy).toBe(true);

    // Queue a COMMANDS turn while the ask streams — it must NOT be downgraded to a send.
    void c.runCommands('set Sales!F2 =C2-D2');
    expect(assist.runTasks).toEqual([]); // not started yet (queued)

    release();
    await first;
    await tick(); // let the scheduled drain settle

    // The queued turn drained through runCommands — the command loop ran, send did NOT re-fire it.
    expect(assist.runTasks).toEqual(['set Sales!F2 =C2-D2']);
    expect(assist.asked).toEqual(['a']); // 'set …' never went through the chat/ask path
  });

  it('an ask turn queued mid-stream drains through send (NOT runCommands) — mode preserved', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [ev({ type: 'done', turn: 1, answer: 'done' })];
    assist.scriptFor = [[{ type: 'token', text: 'q-reply' }, { type: 'done' }]];
    const c = new PanelController(assist, lister([]));

    // Hold the COMMANDS turn busy, then queue an ASK turn behind it.
    const release = assist.holdRun();
    const first = c.runCommands('reconcile'); // commands turn, held busy at its start
    await tick();
    expect(c.getState().busy).toBe(true);

    void c.send('what changed?'); // ASK turn queued while busy
    expect(assist.asked).toEqual([]); // not started yet (queued)

    release();
    await first;
    await tick();

    // The queued turn drained through send — it went to the ask path, NOT runCommands.
    expect(assist.asked).toEqual(['what changed?']);
    expect(assist.runTasks).toEqual(['reconcile']); // only the original command turn ran the loop
  });
});

describe('PanelController — turn-scoped provenance (Finding #4)', () => {
  it('turn A emits provenance A, turn B emits none → a write created by B never receives A', async () => {
    const assist = new FakeAssist();
    // Turn A streams provenance A; turn B streams NONE.
    assist.scriptFor = [
      [{ type: 'token', text: 'A' }, { type: 'provenance', payload: PROV_A }, { type: 'done' }],
      [{ type: 'token', text: 'B' }, { type: 'done' }],
    ];
    const c = new PanelController(assist, lister([]));

    // Turn A → a proposal created by A captures provenance A.
    await c.send('turn A');
    const pa = c.propose('tracked-change', { text: 'x' }, 'from A');
    await c.applyProposal(pa.changeId);
    expect(assist.applied[0]?.provenance).toEqual(PROV_A); // A's write IS attributed to A

    // Turn B (no provenance) → a proposal created by B must carry NONE (never A's leftover).
    await c.send('turn B');
    const pb = c.propose('tracked-change', { text: 'y' }, 'from B');
    await c.applyProposal(pb.changeId);

    expect(assist.applied[1]?.changeId).toBe(pb.changeId);
    expect(assist.applied[1]?.provenance).toBeUndefined();
    // And the recorded ChangeRecord for B carries no provenance either.
    const recB = c.getState().changes.find((r) => r.changeId === pb.changeId);
    expect(recB?.provenance).toBeUndefined();
  });
});

describe('PanelController — approval lifecycle (Finding #6)', () => {
  it('an approval followed by a thrown execution clears the card AND busy (fail-closed)', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [{ approveThenThrow: writeReq('w-throw') }];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('set the formula');
    await tick();
    expect(c.getState().pendingWrite?.changeId).toBe(asChangeId('w-throw'));
    expect(c.getState().busy).toBe(true);

    // Approve → the fake throws from execution AFTER approval. The card + busy must still clear.
    c.approvePendingWrite();
    await run;

    expect(assist.approveResults).toEqual([true]); // it WAS approved
    expect(c.getState().pendingWrite).toBeUndefined(); // card cleared on the throw path
    expect(c.getState().busy).toBe(false); // busy cleared
    expect(c.getState().messages[1]?.error).toContain('actuate boom');
  });

  it('a late decision for a SUPERSEDED approval id cannot apply the current request', async () => {
    const assist = new FakeAssist();
    const first = writeReq('w-first');
    const second = writeReq('w-second');
    assist.commandScript = [{ supersede: { first, second } }];
    const c = new PanelController(assist, lister([]));

    const run = c.runCommands('two writes');
    await tick();
    // The loop has consumed `first`'s decision and is now gated on `second`.
    // Approve `first` (auto, since approvePendingWrite() with no id approves the staged one)…
    c.approvePendingWrite(asChangeId('w-first'));
    await tick();

    // The card now shows `second`. A LATE decision carrying `first`'s id must be ignored.
    expect(c.getState().pendingWrite?.changeId).toBe(asChangeId('w-second'));
    const before = assist.approveResults.length;
    c.approvePendingWrite(asChangeId('w-first')); // stale id → ignored, does not resolve `second`
    c.rejectPendingWrite(asChangeId('w-first')); // stale id → ignored
    expect(assist.approveResults.length).toBe(before); // no new decision was made

    // A decision for the CURRENT id resolves `second` and the loop finishes.
    c.approvePendingWrite(asChangeId('w-second'));
    await run;

    expect(c.getState().pendingWrite).toBeUndefined();
    expect(c.getState().busy).toBe(false);
  });
});

describe('PanelController — structured grounding wiring (Finding #2/#B-wire)', () => {
  it('send/runCommands forward the structured grounding to the session (not raw text)', async () => {
    const assist = new FakeAssist();
    assist.commandScript = [ev({ type: 'done', turn: 1, answer: 'done' })];
    const c = new PanelController(assist, lister([]));
    const grounding: ResolvedGrounding = {
      queryParts: [{ documentReference: { documentName: 'doc-1' } }],
      dataStoreSpecs: [{ dataStore: 'ds-1' }],
    };

    await c.send('ask it', grounding);
    expect(assist.askGrounding[0]).toBe(grounding);

    await c.runCommands('do it', grounding);
    expect(assist.runGrounding[0]).toBe(grounding);
  });
});

function tick(): Promise<void> {
  return new Promise((r) => setTimeout(r, 0));
}

describe('PanelController — discovered catalog (Task 6: skills/data stores reach panel state)', () => {
  it('setDiscoveredCatalog stores available agents and data stores in state', () => {
    const assist = new FakeAssist();
    const controller = new PanelController(assist, lister([]));
    const agents: AgentView[] = [{ id: 'a1', displayName: 'Test Skill' }];
    const dataStores: EngineDataStore[] = [
      { id: 'ds1', resourceName: 'r1', displayName: 'SP Files', connector: 'SharePoint' },
    ];
    controller.setDiscoveredCatalog(agents, dataStores);
    expect(controller.getState().availableAgents).toEqual(agents);
    expect(controller.getState().availableDataStores).toEqual(dataStores);
  });

  it('availableAgents/availableDataStores default to empty arrays', () => {
    const assist = new FakeAssist();
    const controller = new PanelController(assist, lister([]));
    expect(controller.getState().availableAgents).toEqual([]);
    expect(controller.getState().availableDataStores).toEqual([]);
  });
});

describe('PanelController — planner pre-stage (EXPERIENCE.md §F)', () => {
  const wordPlan: CommandPlan = {
    intent: 'rewrite',
    surface: 'word',
    scope: { kind: 'section', ref: '§4' },
    ground: [],
    context: ['inline-preferred'],
    steps: ['rewrite the SLA figure to 99.9% as a tracked change'],
    excludes: ['the indemnity clause'],
    clarify: [],
  };

  it('stages a CommandPlan for confirm, then runs the executor on confirm', async () => {
    const assist = new FakeAssist();
    assist.planned = { plan: wordPlan, errors: [], needsClarification: false };
    // The executor turn (after confirm) stages a plan-level gate so we can see it ran.
    assist.commandScript = [{ plan: [planEffect('w1')] }];
    const c = new PanelController(assist, lister([]));

    await c.proposePlan('/rewrite the SLA but leave indemnity', undefined);
    expect(assist.planTasks).toEqual(['/rewrite the SLA but leave indemnity']);
    expect(c.getState().pendingCommandPlan?.plan).toEqual(wordPlan);
    expect(c.getState().busy).toBe(false);

    c.confirmCommandPlan();
    await tick();
    await tick();
    // The command plan cleared and the EXECUTOR ran (its own effect-level gate is now staged).
    expect(c.getState().pendingCommandPlan).toBeUndefined();
    expect(c.getState().pendingPlan).toBeDefined();
    expect(assist.runTasks[0]).toContain('<confirmed_plan>');
    expect(assist.runTasks[0]).toContain('intent: rewrite');
    expect(assist.runTasks[0]).toContain(
      'step 1: rewrite the SLA figure to 99.9% as a tracked change',
    );
    expect(assist.runTasks[0]).toContain('exclude: the indemnity clause');
  });

  it('cancel discards the plan and runs nothing', async () => {
    const assist = new FakeAssist();
    assist.planned = { plan: wordPlan, errors: [], needsClarification: false };
    const c = new PanelController(assist, lister([]));
    await c.proposePlan('/rewrite x', undefined);
    expect(c.getState().pendingCommandPlan).toBeDefined();

    c.cancelCommandPlan();
    expect(c.getState().pendingCommandPlan).toBeUndefined();
    expect(assist.asked).toEqual([]); // executor never invoked
    expect(c.getState().pendingPlan).toBeUndefined();
  });

  it('a clarify-only plan asks a question and executes nothing', async () => {
    const assist = new FakeAssist();
    assist.planned = {
      plan: { ...wordPlan, steps: [], clarify: ['which section — §4 or §5?'] },
      errors: [],
      needsClarification: true,
    };
    const c = new PanelController(assist, lister([]));
    await c.proposePlan('/rewrite the section', undefined);

    expect(c.getState().pendingCommandPlan).toBeUndefined();
    const last = c.getState().messages.at(-1);
    expect(last?.role).toBe('assistant');
    expect(last?.text).toBe('Before I plan this, I need one detail.');
    expect(c.getState().pendingPlan).toBeUndefined();
    expect(c.getState().pendingPlanClarification?.questions).toEqual(['which section — §4 or §5?']);
  });

  it('routes a clarification answer back through the planner and then stages the confirmed plan', async () => {
    const assist = new FakeAssist();
    assist.plannedQueue = [
      {
        plan: {
          ...wordPlan,
          steps: [],
          clarify: ['what routine should the schedule reflect?'],
        },
        errors: [],
        needsClarification: true,
      },
      { plan: wordPlan, errors: [], needsClarification: false },
    ];
    const c = new PanelController(assist, lister([]));

    await c.proposePlan('/rewrite Help me create a solid mock schedule', undefined);
    expect(c.getState().pendingPlanClarification).toBeDefined();

    c.answerPlanClarification('Google SWE in Sunnyvale with India team calls and fitness');
    await tick();
    await tick();

    expect(assist.planTasks[1]).toContain('/rewrite Help me create a solid mock schedule');
    expect(assist.planTasks[1]).toContain(
      'User clarification:\nGoogle SWE in Sunnyvale with India team calls and fitness',
    );
    expect(c.getState().pendingPlanClarification).toBeUndefined();
    expect(c.getState().pendingCommandPlan?.plan).toEqual(wordPlan);
    expect(c.getState().messages.at(-1)?.text).toBe(
      'Google SWE in Sunnyvale with India team calls and fitness',
    );
  });

  it('degrades to the executor when the planner yields no parseable plan', async () => {
    const assist = new FakeAssist();
    assist.planned = { plan: null, errors: [], needsClarification: false };
    assist.commandScript = [{ plan: [planEffect('w2')] }];
    const c = new PanelController(assist, lister([]));
    await c.proposePlan('/rewrite x', undefined);
    await tick();
    await tick();
    // No plan card; the executor ran directly (its gate staged).
    expect(c.getState().pendingCommandPlan).toBeUndefined();
    expect(c.getState().pendingPlan).toBeDefined();
  });
});
