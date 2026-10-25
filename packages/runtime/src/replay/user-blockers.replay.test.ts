import { describe, expect, it, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
  SseEvent,
} from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession, matchesContextSelector, type CommandLoopEvent } from '../assist-session.js';
import type { DocBridge } from '../bridge.js';

/**
 * Live 2026-10-14: a write to a protected sheet ran 12 turns and ended as a generic error, and a
 * pivot request failed late the same way. Both need the user to act; neither is an error. The task
 * stops at the first one and says why.
 */
const cmd = (program: string) => '```cmd\n' + program + '\n```';

class RefusingBridge implements DocBridge {
  readonly surface = 'excel' as const;
  calls: ActuationRequest[] = [];
  constructor(
    private readonly refusal?: ActuationResult['error'],
    /** Writes that land before the bridge starts refusing. */
    private readonly landFirst = 0,
  ) {}
  getCapabilities(): CapabilityManifest {
    return {
      surface: 'excel',
      contextKinds: ['range'],
      reads: ['read'],
      actuations: [
        { kind: 'write-cells', surface: 'excel', title: 'Write cells', reversible: true },
        { kind: 'format-cells', surface: 'excel', title: 'Format cells', reversible: true },
      ],
    };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.calls.push(request);
    return Promise.resolve(
      this.refusal && this.calls.length > this.landFirst
        ? {
            ok: false,
            changeId: request.changeId,
            kind: request.kind,
            location: 'Sheet2!A1',
            error: this.refusal,
          }
        : { ok: true, changeId: request.changeId, kind: request.kind },
    );
  }
}

function scriptedFetch(turns: string[]) {
  let call = 0;
  const fetchImpl = vi.fn(async () => {
    const text = turns[Math.min(call, turns.length - 1)] ?? cmd('done');
    call += 1;
    const chunk = {
      sessionInfo: { session: 'sess_1' },
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
    };
    return new Response(new TextEncoder().encode(JSON.stringify([chunk])), { status: 200 });
  });
  return { fetch: fetchImpl as unknown as typeof fetch, calls: () => call };
}

async function run(turns: string[], bridge: RefusingBridge, task: string) {
  const { fetch, calls } = scriptedFetch(turns);
  const client = new StreamAssistClient(
    { getAccessToken: () => Promise.resolve('t') },
    { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' },
    fetch,
  );
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'excel' as const } },
    context: { docState: false },
  });
  const events: Array<SseEvent | CommandLoopEvent> = [];
  let thrown: unknown;
  try {
    for await (const e of session.runCommands(task, { approvePlan: () => true })) events.push(e);
  } catch (error) {
    thrown = error;
  }
  return { events, thrown, modelCalls: calls(), runs: session.executions.list() };
}

const write = cmd('set Sheet2!A1 Region');

describe('a refusal the user must fix ends the task at once', () => {
  it('stops after the first host refusal, without an error', async () => {
    const bridge = new RefusingBridge({
      code: 'host_refused',
      reason: 'protected',
      message: 'Sheet2 is protected.',
    });
    const { events, thrown, modelCalls, runs } = await run(
      [write, write, write],
      bridge,
      'Put Region in A1',
    );
    expect(thrown).toBeUndefined();
    expect(modelCalls).toBe(1);
    expect(bridge.calls).toHaveLength(1);
    expect(events.filter((e) => e.type === 'needs-user')).toEqual([
      {
        type: 'needs-user',
        turn: 1,
        blocker: {
          cause: 'host_refused',
          reason: 'protected',
          kind: 'write-cells',
          location: 'Sheet2!A1',
        },
      },
    ]);
    expect(runs.at(-1)?.status).toBe('needs-user');
  });

  it('a refused verified block stops plainly, without the "could not be confirmed" error (P5)', async () => {
    // Argon mx-xl-refused-02 (2026-10-19): `set` ×2 + `finish when=verified` on a protected sheet.
    const bridge = new RefusingBridge({
      code: 'host_refused',
      reason: 'protected',
      message: 'Sheet2 is protected.',
    });
    const { events } = await run(
      [cmd('set Sheet2!A1 "Remarks"\nset Sheet2!A2 "review"\nfinish when=verified')],
      bridge,
      'Put Remarks in A1',
    );
    expect(events.filter((e) => e.type === 'error')).toEqual([]);
    expect(events.filter((e) => e.type === 'needs-user')).toHaveLength(1);
  });

  it('keeps "incomplete" when an earlier write landed, so asking again cannot repeat it', async () => {
    const bridge = new RefusingBridge(
      { code: 'host_refused', reason: 'protected', message: 'Sheet2 is protected.' },
      1,
    );
    const { events, modelCalls, runs } = await run(
      [cmd('set Sheet1!A1 Region\nset Sheet2!A1 Region')],
      bridge,
      'Put Region in A1 of both sheets',
    );
    expect(modelCalls).toBe(1);
    expect(events.some((e) => e.type === 'needs-user')).toBe(true);
    expect(runs.at(-1)?.status).toBe('incomplete');
  });

  it('stops at an unsupported command before anything is staged', async () => {
    const bridge = new RefusingBridge();
    const { events, thrown, modelCalls, runs } = await run(
      [
        cmd(
          '/insert-pivot sourceRange=Sheet1!A1:C9 destinationRange=Sheet1!E1 rowFields=Region valueFields=Revenue',
        ),
        cmd('set Sheet1!E1 Region'),
      ],
      bridge,
      'Make a pivot table',
    );
    expect(thrown).toBeUndefined();
    expect(modelCalls).toBe(1);
    expect(bridge.calls).toHaveLength(0);
    expect(events.filter((e) => e.type === 'needs-user')).toEqual([
      {
        type: 'needs-user',
        turn: 1,
        blocker: {
          cause: 'not_supported',
          kind: 'insert-pivot',
          available: ['write-cells', 'format-cells'],
        },
      },
    ]);
    expect(runs.at(-1)?.status).toBe('needs-user');
  });

  it('stops a repeated program even when its suggested_questions block changes each turn', async () => {
    // The engine adds a fresh follow-up block every turn; it must not hide the repeat (2026-10-19).
    const bridge = new RefusingBridge();
    const looping = (n: number) =>
      cmd('help discover flag') + `\n\n\`\`\`suggested_questions\n["Question ${n}?"]\n\`\`\``;
    const { events, modelCalls, thrown } = await run(
      [looping(1), looping(2), looping(3), looping(4), looping(5)],
      bridge,
      'Flag this for follow-up',
    );
    expect(thrown).toBeUndefined();
    expect(modelCalls).toBe(3);
    expect(events.filter((e) => e.type === 'needs-user')).toEqual([
      { type: 'needs-user', turn: 3, blocker: { cause: 'no_progress' } },
    ]);
  });

  it('stops when the same read-only program comes back a third time', async () => {
    // Live 2026-10-15 (mx-ol-unsupported-03): `properties` repeated until turn 12.
    const bridge = new RefusingBridge();
    const looping = cmd('help discover flag');
    const { events, modelCalls, runs, thrown } = await run(
      [looping, cmd('help set'), looping, looping, looping, looping],
      bridge,
      'Flag this for follow-up',
    );
    expect(thrown).toBeUndefined();
    expect(modelCalls).toBe(4);
    expect(events.filter((e) => e.type === 'needs-user')).toEqual([
      { type: 'needs-user', turn: 4, blocker: { cause: 'no_progress' } },
    ]);
    expect(runs.at(-1)?.status).toBe('needs-user');
  });

  it('does not count a repeated read across turns that land changes', async () => {
    const bridge = new RefusingBridge();
    const check = cmd('help set');
    const { events } = await run(
      [
        check,
        cmd('set Sheet1!A1 1'),
        check,
        cmd('set Sheet1!A2 2'),
        check,
        cmd('set Sheet1!A3 3'),
        cmd('done'),
      ],
      bridge,
      'Fill A1:A3',
    );
    expect(events.some((e) => e.type === 'needs-user')).toBe(false);
    expect(bridge.calls).toHaveLength(3);
  });

  it('a clean rejection followed by the model explaining why ends as an answer, not an error', async () => {
    // SIT 2026-10-14 #4: "change the subject" with a received mail open ended on the error card
    // even though the model explained the limit.
    const bridge = new RefusingBridge({
      code: 'no_compose',
      message: 'The open item is a received message, not a draft being written.',
    });
    const { events, thrown, runs } = await run(
      [cmd('set Sheet2!A1 Region'), cmd('done "A received email cannot be changed."')],
      bridge,
      'Change the subject',
    );
    expect(thrown).toBeUndefined();
    expect(runs.at(-1)?.status).toBe('needs-user');
    expect(events.filter((e) => e.type === 'done' && 'turn' in e).at(-1)).toMatchObject({
      message: 'A received email cannot be changed.',
    });
  });

  it('sends a core verb used on the wrong surface back for repair instead of stopping', async () => {
    // Live 2026-10-15 (mx-ol-read-10): `reply` means comment-reply, which this surface lacks;
    // the model meant something else and corrects it on the repair turn.
    const bridge = new RefusingBridge();
    const { events, modelCalls, runs } = await run(
      [cmd('reply c1 "Thanks"\ndone'), cmd('set Sheet1!E1 Region'), cmd('done')],
      bridge,
      'Answer the comment',
    );
    expect(events.some((e) => e.type === 'needs-user')).toBe(false);
    expect(events.some((e) => e.type === 'repair')).toBe(true);
    expect(modelCalls).toBeGreaterThan(1);
    expect(runs.at(-1)?.status).not.toBe('needs-user');
  });
});

describe('a read/help-only loop ends in an honest "not supported yet" (WS4a)', () => {
  // Argon mx-wd-table-14 / mx-xl-condfmt-08: read ↔ help, alternating, until turn 12.
  const read = cmd('help format');
  const help = cmd('help set');

  it('asks once for a plain answer, then stops as not supported', async () => {
    const bridge = new RefusingBridge();
    const { events, modelCalls } = await run([read, help, read, help, read], bridge, 'x');
    expect(modelCalls).toBe(4);
    expect(events.filter((e) => e.type === 'needs-user')).toEqual([
      {
        type: 'needs-user',
        turn: 4,
        blocker: { cause: 'not_supported', available: ['write-cells', 'format-cells'] },
      },
    ]);
  });

  it('after a landed write, a stale loop ends as "could not finish", not "not supported" (review L3)', async () => {
    const bridge = new RefusingBridge();
    const { events } = await run(
      [cmd('set Sheet2!A1 Region'), read, help, read, help],
      bridge,
      'x',
    );
    expect(events.filter((e) => e.type === 'needs-user')).toEqual([
      { type: 'needs-user', turn: 5, blocker: { cause: 'no_progress' } },
    ]);
  });

  it('the nudge reaches the model after the first stale turn', async () => {
    const bridge = new RefusingBridge();
    const { fetch } = scriptedFetch([read, help, read, cmd('done "Cannot do that yet."')]);
    const client = new StreamAssistClient(
      { getAccessToken: () => Promise.resolve('t') },
      { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' },
      fetch,
    );
    const session = new AssistSession(bridge, client, {
      unit: { connectors: [], surfaceContext: { kind: 'excel' as const } },
      context: { docState: false },
    });
    for await (const _ of session.runCommands('x', { approvePlan: () => true })) void _;
    const bodies = (
      fetch as unknown as { mock: { calls: Array<[unknown, RequestInit]> } }
    ).mock.calls.map(([, init]) => String(init.body));
    expect(bodies[2]).not.toContain('cannot be done with the listed commands');
    expect(bodies[3]).toContain('cannot be done with the listed commands');
  });
});

describe('a read never comes back silently empty (WS3, Argon mx-xl-unsupported-05)', () => {
  it('names the selector when nothing resolves', async () => {
    const bridge = new RefusingBridge();
    (bridge as unknown as { readRange: () => Promise<ResolvedContext[]> }).readRange = () =>
      Promise.resolve([]);
    const { events } = await run([cmd('read NoSuch!A1:F21'), cmd('done "x"')], bridge, 'read it');
    const read = events.find((e) => e.type === 'read-result') as { result?: { error?: string } };
    expect(read.result?.error).toMatch(/NoSuch!A1:F21/);
    expect(read.result?.error).toMatch(/'My Sheet'!A1/);
  });
});

describe('address selectors match however the sheet name is quoted', () => {
  const ref = {
    id: 'xl:comment:c2',
    kind: 'comment',
    surface: 'excel',
    title: 'Comment on D2',
    live: false,
    anchor: { locator: "'Branch Performance'!D2" },
  } as unknown as Parameters<typeof matchesContextSelector>[0];

  it.each([
    "'Branch Performance'!D2",
    'Branch Performance!D2',
    '"Branch Performance!D2"',
    "'Branch Performance'!$D$2",
    'branch performance!d2',
  ])('%s', (selector) => {
    expect(matchesContextSelector(ref, selector)).toBe(true);
  });

  it('still tells different cells apart', () => {
    expect(matchesContextSelector(ref, 'Branch Performance!D3')).toBe(false);
  });
});
