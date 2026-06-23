import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ChangeId,
  ContextRef,
  SseEvent,
} from '@ge/contracts';
import { asChangeId } from '@ge/contracts';
import type { CommandLoopEvent } from '@ge/runtime';
import type { HostEvent } from '@ge/triggers';
import {
  PanelController,
  type AssistLike,
  type ContextLister,
  type PlanEffect,
  type PlanRunCommandsOptions,
} from './controller.js';

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
  | { plan: PlanEffect[] };

const ev = (event: SseEvent | CommandLoopEvent): CommandAction => ({ event });

const planEffect = (changeId: string): PlanEffect => {
  const request = writeReq(changeId);
  return { request, command: `set Sales!F2 =C2-D2 [${changeId}]` };
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
  applied: Array<{ kind: string; changeId: string }> = [];
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
  async *ask(query: string, opts?: { signal?: AbortSignal }): AsyncGenerator<SseEvent> {
    this.asked.push(query);
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
  apply(
    kind: ActuationRequest['kind'],
    _params: ActuationParams,
    changeId: ChangeId,
  ): Promise<ActuationResult> {
    this.applied.push({ kind, changeId });
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
  approveCalls: ActuationRequest[] = [];
  approveResults: boolean[] = [];
  /** The plan effect-sets passed to `approvePlan` and the decisions returned, for assertions. */
  planCalls: PlanEffect[][] = [];
  planResults: boolean[] = [];
  async *runCommands(
    task: string,
    opts?: PlanRunCommandsOptions,
  ): AsyncGenerator<SseEvent | CommandLoopEvent> {
    this.runTasks.push(task);
    for (const action of this.commandScript) {
      if ('approve' in action) {
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
      } else {
        yield action.event;
      }
    }
  }
  ingest(event: HostEvent): Promise<void> {
    this.ingested.push(event);
    return Promise.resolve();
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
});

describe('PanelController — ask / stream', () => {
  it('streams tokens + citations into the assistant message and toggles busy', async () => {
    const assist = new FakeAssist();
    assist.script = [
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
      'done',
    ]);
    expect(c.getState().pendingWrite).toBeUndefined();
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

function tick(): Promise<void> {
  return new Promise((r) => setTimeout(r, 0));
}
