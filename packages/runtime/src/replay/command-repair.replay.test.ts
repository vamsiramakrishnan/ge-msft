import { readFileSync } from 'node:fs';
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
import { AssistSession, type CommandLoopEvent } from '../assist-session.js';
import type { DocBridge } from '../bridge.js';

/**
 * Fix D (docs/COMMAND-RELIABILITY.md) replayed against the executor programs Gemini produced live on
 * 2026-09-30. Runs r2 and r4 each emitted an invalid line followed by `done`; before the fix the task
 * ended as incomplete without the model ever seeing its error.
 */
const live = JSON.parse(
  readFileSync(new URL('./live-2026-09-30-excel-chart.json', import.meta.url), 'utf8'),
) as { executorPrograms: Record<string, string> };

const cmd = (program: string) => '```cmd\n' + program + '\n```';

class ExcelLikeBridge implements DocBridge {
  readonly surface = 'excel' as const;
  applied: ActuationRequest[] = [];
  getCapabilities(): CapabilityManifest {
    return {
      surface: 'excel',
      contextKinds: ['range'],
      reads: ['read'],
      actuations: [
        { kind: 'write-cells', surface: 'excel', title: 'Write cells', reversible: true },
        { kind: 'insert-chart', surface: 'excel', title: 'Insert chart', reversible: true },
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
    this.applied.push(request);
    return Promise.resolve({ ok: true, changeId: request.changeId, kind: request.kind });
  }
  readRange(a1: string): Promise<ResolvedContext[]> {
    const gfm = '| Product | Total |\n| --- | --- |\n| Laptop | 130000 |\n| Mouse | 12000 |';
    return Promise.resolve([
      {
        ref: { id: `xl:${a1}`, kind: 'range', surface: 'excel', title: a1, live: false },
        value: { as: 'text', text: gfm, mimeType: 'text/markdown' },
      },
    ]);
  }
}

function streamOf(pieces: string[]): ReadableStream<Uint8Array> {
  const enc = new TextEncoder();
  let i = 0;
  return new ReadableStream({
    pull(c) {
      if (i < pieces.length) c.enqueue(enc.encode(pieces[i++]!));
      else c.close();
    },
  });
}

/** A streamAssist fetch that replays one scripted model reply per call and records each query. */
function scriptedFetch(turns: string[]) {
  const queries: string[] = [];
  let call = 0;
  const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
    const body = JSON.parse(init?.body ?? '{}') as {
      query?: { text?: string; parts?: Array<{ text?: string }> };
    };
    queries.push(body.query?.text ?? (body.query?.parts ?? []).map((p) => p.text ?? '').join('\n'));
    const text = turns[Math.min(call, turns.length - 1)] ?? cmd('done');
    call += 1;
    const chunk = {
      sessionInfo: { session: 'sess_1' },
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
    };
    return new Response(streamOf([JSON.stringify([chunk])]), { status: 200 });
  });
  return { fetch: fetchImpl as unknown as typeof fetch, queries };
}

type SessionOptions = ConstructorParameters<typeof AssistSession>[2];
type LoopOptions = NonNullable<Parameters<AssistSession['runCommands']>[1]>;

function setup(turns: string[], extra: Partial<SessionOptions> = {}) {
  const bridge = new ExcelLikeBridge();
  const { fetch, queries } = scriptedFetch(turns);
  const client = new StreamAssistClient(
    { getAccessToken: () => Promise.resolve('t') },
    { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' },
    fetch,
  );
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'excel' as const } },
    context: { docState: false },
    ...extra,
  });
  return { bridge, queries, session };
}

async function collect(gen: AsyncGenerator<SseEvent | CommandLoopEvent>) {
  const events: Array<SseEvent | CommandLoopEvent> = [];
  for await (const e of gen) events.push(e);
  return events;
}

async function run(
  turns: string[],
  extra: Partial<SessionOptions> = {},
  loop: Partial<LoopOptions> = {},
) {
  const { bridge, queries, session } = setup(turns, extra);
  const events = await collect(
    session.runCommands('Create a bar chart of Total by Product', {
      approvePlan: () => true,
      ...loop,
    }),
  );
  return { bridge, events, queries };
}

function fakeSharedStore() {
  const shared = new Map<string, string>();
  return {
    shared,
    sharedStore: {
      list: () =>
        Promise.resolve([...shared.entries()].map(([name, text]) => ({ name, size: text.length }))),
      read: (path: string) => Promise.resolve(shared.get(path)),
      write: (path: string, content: string) => {
        shared.set(path, content);
        return Promise.resolve();
      },
      remove: (path: string) => {
        shared.delete(path);
        return Promise.resolve();
      },
    },
  };
}

const ofType = <T extends (SseEvent | CommandLoopEvent)['type']>(
  events: Array<SseEvent | CommandLoopEvent>,
  type: T,
) => events.filter((e) => e.type === type);

describe('D — an invalid program is sent back for repair instead of ending the task', () => {
  it.each(['r2', 'r4'])(
    'recorded %s: done is ignored and the model gets its errors',
    async (id) => {
      const { bridge, events, queries } = await run([
        cmd(live.executorPrograms[id]!),
        cmd('chart bar Sheet2!A1:J11 title="Total by Product"'),
        cmd('done'),
      ]);
      // Turn 1 staged and applied nothing, and did not finish the task.
      const repair = ofType(events, 'repair');
      expect(repair).toHaveLength(1);
      expect(repair[0]).toMatchObject({ turn: 1 });
      expect(events.some((e) => e.type === 'done' && 'turn' in e && e.turn === 1)).toBe(false);
      // The next query carried the errors and the repair instruction to the model.
      expect(queries[1]).toContain('Program not applied');
      // The corrected program then flowed through the normal gate: one chart, then done.
      expect(bridge.applied.map((r) => r.kind)).toEqual(['insert-chart']);
      expect(events.some((e) => e.type === 'done' && 'turn' in e && e.turn === 3)).toBe(true);
    },
  );

  it('refuses the recorded r1 JSON chart source before approval and repairs it (fix C)', async () => {
    // Live r1 emitted `chart bar {"source":"Sheet2!A1:J11",…}`; it reached the approval card and
    // would only have failed inside Excel after the user approved it.
    const { bridge, events, queries } = await run([
      cmd(
        live.executorPrograms
          .r1!.split('\n')
          .filter((l) => l.startsWith('chart'))
          .join('\n') + '\ndone',
      ),
      cmd('chart bar Sheet2!C1:C11,Sheet2!G1:G11 title="Total by Product"'),
      cmd('done'),
    ]);
    expect(ofType(events, 'plan-preview')[0]).toMatchObject({ turn: 2 });
    expect(queries[1]).toContain('chart range must be A1 areas');
    expect(bridge.applied.map((r) => r.params.chart?.sourceRange)).toEqual([
      'Sheet2!C1:C11,Sheet2!G1:G11',
    ]);
  });

  it('does not accept done in the same block as a read (live w2: `read` + `done`)', async () => {
    const { bridge, events, queries } = await run([
      cmd('read Sheet2!A1:B3\ndone'),
      cmd('set Sheet2!L1 "x"'),
      cmd('done'),
    ]);
    expect(events.some((e) => e.type === 'done' && 'turn' in e && e.turn === 1)).toBe(false);
    expect(queries[1]).toContain('done cannot be batched with a read');
    expect(bridge.applied).toHaveLength(1);
  });

  it('keeps a new write when the same program repeats one that already landed (24a4e81 stays intact)', async () => {
    // Regression review: the duplicate-write refusal from 24a4e81 was counted as a failed line, so
    // the new `set Sheet2!A2` was withheld and the task ended repair_exhausted.
    const { bridge, events } = await run([
      cmd('set Sheet2!A1 "1"'),
      cmd('set Sheet2!A1 "1"\nset Sheet2!A2 "2"'),
      cmd('done'),
    ]);
    expect(bridge.applied.map((r) => r.params.target?.range)).toEqual(['Sheet2!A1', 'Sheet2!A2']);
    expect(ofType(events, 'repair')).toHaveLength(0);
    expect(events.some((e) => e.type === 'done')).toBe(true);
  });

  it('does not end as repair_exhausted when the model repeats a landed write', async () => {
    const again = cmd('set Sheet2!A1 "1"');
    const { bridge, events } = await run([again, again, again, again, cmd('done')]);
    expect(bridge.applied).toHaveLength(1); // the duplicate is still refused, never applied twice
    expect(events.some((e) => e.type === 'error' && e.code === 'repair_exhausted')).toBe(false);
    expect(ofType(events, 'repair')).toHaveLength(0);
  });

  it('withholds a valid write that shares a program with a failed line', async () => {
    const { bridge, events } = await run([
      cmd('set Sheet2!L1 "Total"\nanalyze {"action":"chart","type":"bar"}\ndone'),
      cmd('set Sheet2!L1 "Total"'),
      cmd('done'),
    ]);
    // Applying the valid `set` from turn 1 and again from the corrected program would double-write.
    expect(ofType(events, 'repair')[0]).toMatchObject({ withheldWrites: 1 });
    expect(bridge.applied).toHaveLength(1);
  });

  it('stops as incomplete after the repair budget, having applied nothing', async () => {
    const bad = cmd(live.executorPrograms.r2!);
    const { bridge, events } = await run([bad, bad, bad, bad]);
    expect(ofType(events, 'repair')).toHaveLength(3);
    expect(events.some((e) => e.type === 'error' && e.code === 'repair_exhausted')).toBe(true);
    expect(ofType(events, 'plan-preview')).toHaveLength(0);
    expect(bridge.applied).toHaveLength(0);
  });

  it('names a share that already landed so the corrected program does not repeat it', async () => {
    // Security review, medium 1: an approved share lands inline and cannot be withheld.
    const { sharedStore, shared } = fakeSharedStore();
    const { bridge, events, queries } = await run(
      [
        cmd(
          'share out.txt = read Sheet2!A1:B3\nset Sheet2!L1 "x"\nanalyze {"action":"chart"}\ndone',
        ),
        cmd('set Sheet2!L1 "x"'),
        cmd('done'),
      ],
      { sharedStore, estateWritesEnabled: true },
      { approveShare: () => true },
    );
    expect(ofType(events, 'repair')[0]).toMatchObject({
      withheldWrites: 1,
      alreadyApplied: ['share out.txt'],
    });
    expect(queries[1]).toContain('Already applied and must not be repeated: share out.txt');
    expect([...shared.keys()].filter((k) => k === 'out.txt')).toHaveLength(1);
    expect(bridge.applied).toHaveLength(1);
  });

  it('treats a declined share as a decision, not a failed line', async () => {
    // Security review, low 3: a denial must not be re-asked through a repair turn.
    const { sharedStore } = fakeSharedStore();
    const { bridge, events } = await run(
      [cmd('share out.txt = read Sheet2!A1:B3\nset Sheet2!L1 "x"'), cmd('done')],
      { sharedStore, estateWritesEnabled: true },
      { approveShare: () => false },
    );
    expect(ofType(events, 'repair')).toHaveLength(0);
    expect(bridge.applied).toHaveLength(1);
  });

  it('discards the let bindings of a repaired program', async () => {
    // Security review, medium 2 and live run a-r2: a surviving binding made the corrected
    // program fail on its own `let`.
    const { bridge, queries } = await run([
      cmd('let $x = read Sheet2!A1:B3\nanalyze {"action":"chart"}\ndone'),
      cmd('set Sheet2!L1 = ($x | count)'),
      cmd('let $x = read Sheet2!A1:B3\nset Sheet2!L1 = ($x | count)'),
      cmd('done'),
    ]);
    expect(queries[2]).toContain('unbound variable');
    expect(bridge.applied).toHaveLength(1);
  });

  it('stops after consecutive failing turns even when they stage no write', async () => {
    // Security review, low 4 and live run a-r2: repeated failures without a write or `done` ran
    // all 12 turns instead of stopping.
    const bad = cmd('read Sheet2!A1:B3\nfoo bar');
    const { events, queries } = await run([bad, bad, bad, bad, bad]);
    expect(events.some((e) => e.type === 'error' && e.code === 'repair_exhausted')).toBe(true);
    expect(queries).toHaveLength(3);
  });

  it('sends a verified program back for repair too, applying none of it', async () => {
    const { bridge, events } = await run([
      cmd('set Sheet2!L1 "x"\nanalyze {"action":"chart"}\nfinish when=verified'),
      cmd('done'),
    ]);
    expect(ofType(events, 'repair')).toHaveLength(1);
    expect(bridge.applied).toHaveLength(0);
  });

  it('does not repair a direct program, which has no model to fix it', async () => {
    const { bridge, session } = setup([]);
    const events = await collect(
      session.runCommandProgram('set Sheet2!L1 "x"\nfoo bar', { approvePlan: () => true }),
    );
    expect(ofType(events, 'repair')).toHaveLength(0);
    expect(bridge.applied).toHaveLength(1);
  });
});
