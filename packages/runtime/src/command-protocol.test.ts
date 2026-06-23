import { describe, it, expect } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
  SseEvent,
} from '@ge/contracts';
import { ActuationRequestSchema } from '@ge/contracts';
import type { StreamAssistClient, StreamOptions } from '@ge/gemini-client';
import type { AssistRequest } from '@ge/contracts';
import { TriggerRegistry } from '@ge/triggers';
import { AssistSession, type CommandLoopEvent } from './assist-session.js';
import { compileCommand, renderGrammarPrompt } from './command-protocol.js';
import type { DocBridge } from './bridge.js';
import { asChangeId } from '@ge/contracts';

/* ───────────────────────── compileCommand unit tests ──────────────────── */

const mint = () => asChangeId('cid-fixed');

describe('compileCommand', () => {
  it('compiles `set` → a valid write-cells ActuationRequest', () => {
    const c = compileCommand(
      { verb: 'set', cell: 'Sales!F2', value: '=C2-D2' },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'write-cells',
        surface: 'excel',
        params: { target: { range: 'Sales!F2' }, cells: [['=C2-D2']] },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `suggest` → a tracked-change with target.matchText', () => {
    const c = compileCommand(
      { verb: 'suggest', oldText: 'old', newText: 'new' },
      { surface: 'word', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: { kind: 'tracked-change', params: { target: { matchText: 'old' }, text: 'new' } },
    });
  });

  it('compiles `comment` → add-comment with a cell range target (Excel)', () => {
    const c = compileCommand(
      { verb: 'comment', selector: 'Sales!A16', text: 'anomalous spike' },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'add-comment',
        surface: 'excel',
        params: { target: { range: 'Sales!A16' }, text: 'anomalous spike' },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `comment` → add-comment with a matchText anchor (Word)', () => {
    const c = compileCommand(
      { verb: 'comment', selector: 'the SLA is 99.5%', text: 'needs a source' },
      { surface: 'word', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'add-comment',
        surface: 'word',
        params: { target: { matchText: 'the SLA is 99.5%' }, text: 'needs a source' },
      },
    });
  });

  it('compiles `format` → format-cells with typed format params', () => {
    const c = compileCommand(
      {
        verb: 'format',
        range: 'Sales!A16:C16',
        props: { bold: 'true', italic: 'false', fill: '#FFF2CC', numberFormat: '$#,##0.00' },
      },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'format-cells',
        surface: 'excel',
        params: {
          target: { range: 'Sales!A16:C16' },
          format: { bold: true, italic: false, fill: '#FFF2CC', numberFormat: '$#,##0.00' },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('ignores unknown format keys but errors when NO recognized prop is present', () => {
    const ok = compileCommand(
      { verb: 'format', range: 'A1', props: { bold: 'true', wibble: 'x' } },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(ok).toMatchObject({ kind: 'write', request: { params: { format: { bold: true } } } });

    const bad = compileCommand(
      { verb: 'format', range: 'A1', props: { wibble: 'x' } },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(bad).toMatchObject({ error: expect.stringContaining('recognized property') });
  });

  it('compiles `reply` → comment-reply with a commentId target (Zod-valid, changeId minted once)', () => {
    const ids: string[] = [];
    const mintOnce = () => {
      const id = asChangeId(`cid-${ids.length}`);
      ids.push(id);
      return id;
    };
    const c = compileCommand(
      { verb: 'reply', commentId: '{3f2a}', text: 'addressed in the redline' },
      { surface: 'word', mintChangeId: mintOnce },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'comment-reply',
        surface: 'word',
        params: { target: { commentId: '{3f2a}' }, text: 'addressed in the redline' },
      },
    });
    expect(ids).toHaveLength(1); // changeId minted exactly once
    if ('request' in c) {
      expect(c.request.changeId).toBe('cid-0');
      expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
    }
  });

  it('compiles reads to read intents', () => {
    expect(compileCommand({ verb: 'outline' }, { surface: 'excel', mintChangeId: mint })).toEqual({
      kind: 'read',
      intent: { read: 'outline' },
    });
    expect(
      compileCommand({ verb: 'read', selector: 'A1:B2' }, { surface: 'excel', mintChangeId: mint }),
    ).toEqual({ kind: 'read', intent: { read: 'range', selector: 'A1:B2' } });
    expect(
      compileCommand({ verb: 'search', text: 'x' }, { surface: 'word', mintChangeId: mint }),
    ).toEqual({ kind: 'read', intent: { read: 'search', text: 'x' } });
  });

  it('compiles control verbs', () => {
    expect(compileCommand({ verb: 'done' }, { surface: 'excel', mintChangeId: mint })).toEqual({
      kind: 'control',
      verb: 'done',
    });
  });
});

describe('renderGrammarPrompt', () => {
  it('advertises set for Excel and not suggest', () => {
    const prompt = renderGrammarPrompt(excelManifest);
    expect(prompt).toContain('set <A1> <value|=formula>');
    expect(prompt).toContain('read <A1|NamedRange>');
    expect(prompt).not.toContain('suggest "old text"');
    expect(prompt).toContain('```cmd');
  });

  it('advertises suggest for Word and not set', () => {
    const prompt = renderGrammarPrompt(wordManifest);
    expect(prompt).toContain('suggest "old text" => "new text"');
    expect(prompt).not.toContain('set <A1');
  });
});

/* ───────────────────────── loop fixtures ──────────────────────────────── */

const excelManifest: CapabilityManifest = {
  surface: 'excel',
  contextKinds: ['range', 'sheet'],
  reads: ['outline', 'read', 'search'],
  actuations: [{ kind: 'write-cells', surface: 'excel', title: 'Write cells', reversible: true }],
};

const wordManifest: CapabilityManifest = {
  surface: 'word',
  contextKinds: ['selection', 'document'],
  reads: ['outline', 'read', 'search'],
  actuations: [
    { kind: 'tracked-change', surface: 'word', title: 'Insert tracked change', reversible: true },
  ],
};

function snapshot(surface: 'excel' | 'word', version: number): DocStateSnapshot {
  return {
    surface,
    version,
    capturedAt: '2026-06-22T00:00:00Z',
    outline: [],
    inventory: [],
  };
}

/** A fake Excel bridge recording actuations, reads, and serving a versioned doc-state. */
class FakeExcelBridge implements DocBridge {
  readonly surface = 'excel' as const;
  applied: ActuationRequest[] = [];
  reads: string[] = [];
  version = 1;

  getCapabilities(): CapabilityManifest {
    return excelManifest;
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.applied.push(request);
    this.version += 1;
    return Promise.resolve({
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      location: 'F2',
    });
  }
  captureDocState(): Promise<DocStateSnapshot | undefined> {
    return Promise.resolve(snapshot('excel', this.version));
  }
  readRange(a1: string): Promise<ResolvedContext[]> {
    this.reads.push(a1);
    return Promise.resolve([
      {
        ref: { id: `xl:${a1}`, kind: 'range', surface: 'excel', title: a1, live: false },
        value: { as: 'text', text: `values of ${a1}` },
      },
    ]);
  }
  searchDocument(query: string): Promise<ResolvedContext[]> {
    this.reads.push(`search:${query}`);
    return Promise.resolve([
      {
        ref: { id: `xl:search`, kind: 'range', surface: 'excel', title: query, live: false },
        value: { as: 'text', text: `rows matching ${query}` },
      },
    ]);
  }
}

class FakeWordBridge implements DocBridge {
  readonly surface = 'word' as const;
  applied: ActuationRequest[] = [];
  getCapabilities(): CapabilityManifest {
    return wordManifest;
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.applied.push(request);
    return Promise.resolve({
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      location: 'para:1',
    });
  }
  captureDocState(): Promise<DocStateSnapshot | undefined> {
    return Promise.resolve(snapshot('word', 1));
  }
  searchDocument(): Promise<ResolvedContext[]> {
    return Promise.resolve([
      {
        ref: { id: 'w:1', kind: 'paragraph', surface: 'word', title: 'p1', live: false },
        value: { as: 'text', text: 'The SLA is 99.5%.' },
      },
    ]);
  }
}

/**
 * A fake StreamAssistClient that replays a scripted transcript: one string of answer text per
 * model turn (wrapped in token + provenance + done SSE events). Records the queries it received.
 */
function fakeClient(turns: string[]): {
  client: StreamAssistClient;
  queries: string[];
} {
  const queries: string[] = [];
  let i = 0;
  const stream = async function* (
    req: AssistRequest,
    _opts: StreamOptions,
  ): AsyncGenerator<SseEvent> {
    queries.push(req.query ?? '');
    const text = turns[i++] ?? '```cmd\ndone\n```';
    yield { type: 'token', text };
    yield {
      type: 'provenance',
      payload: {
        agentId: 'gemini-enterprise:e',
        identity: 'v.k@acme',
        timestamp: '2026-06-22T00:00:00Z',
        sources: [],
        contentHash: 'h',
        sessionId: 'sess_loop' as never,
      },
    };
    yield { type: 'done' };
  };
  const client = { stream } as unknown as StreamAssistClient;
  return { client, queries };
}

const unit = { connectors: [], surfaceContext: { kind: 'excel' as const } };

async function collect(
  gen: AsyncGenerator<SseEvent | CommandLoopEvent>,
): Promise<Array<SseEvent | CommandLoopEvent>> {
  const out: Array<SseEvent | CommandLoopEvent> = [];
  for await (const e of gen) out.push(e);
  return out;
}

function loopEvents(events: Array<SseEvent | CommandLoopEvent>): CommandLoopEvent[] {
  const kinds = new Set([
    'turn-start',
    'command',
    'read-result',
    'write-result',
    'no-fence',
    'capped',
    'done',
    'exhausted',
  ]);
  return events.filter((e) => kinds.has(e.type)) as CommandLoopEvent[];
}

/* ───────────────────────── the loop ───────────────────────────────────── */

describe('AssistSession.runCommands — the bounded command loop', () => {
  it('read-many: batches all reads in a turn, then terminates on done', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      '**thought** discover first\n```cmd\noutline\nread Sales!C2:C7\nsearch margin\n```',
      '**answer** all set\n```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('Analyze the sheet'));
    const loop = loopEvents(events);

    // Three reads executed in turn 1.
    const reads = loop.filter((e) => e.type === 'read-result');
    expect(reads).toHaveLength(3);
    expect(bridge.reads).toEqual(['Sales!C2:C7', 'search:margin']); // outline uses captureDocState
    // Terminated on done.
    expect(loop.at(-1)).toMatchObject({ type: 'done' });
    // Turn 1 query carries the protocol + the task.
    expect(queries[0]).toContain('TASK:');
    expect(queries[0]).toContain('Analyze the sheet');
    // Turn 2 query is the ```result block fed back.
    expect(queries[1]).toContain('```result');
  });

  it('write-one: a set compiles to a gated write-cells request, one at a time', async () => {
    const bridge = new FakeExcelBridge();
    const gate = new TriggerRegistry();
    const seen: ActuationRequest[] = [];
    gate.register({
      id: 'audit',
      on: 'pre-actuation',
      handle: (e) => {
        if (e.type === 'pre-actuation') seen.push(e.request);
        return { kind: 'continue' };
      },
    });
    const { client } = fakeClient([
      '```cmd\nset Sales!F2 =SUM(C2:C7)\nset Sales!F3 =SUM(D2:D7)\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit, triggers: gate });

    const events = await collect(session.runCommands('Write totals', { approveWrite: () => true }));
    const writes = loopEvents(events).filter((e) => e.type === 'write-result');

    expect(writes).toHaveLength(2);
    expect(bridge.applied).toHaveLength(2);
    expect(seen).toHaveLength(2); // each gated
    expect(bridge.applied[0]).toMatchObject({
      kind: 'write-cells',
      surface: 'excel',
      params: { target: { range: 'Sales!F2' }, cells: [['=SUM(C2:C7)']] },
    });
    // Provenance stamped from the streamed turn.
    expect(bridge.applied[0]!.provenance?.identity).toBe('v.k@acme');
  });

  it('a blocked gate yields a corrective write-result, not a thrown loop', async () => {
    const bridge = new FakeExcelBridge();
    const gate = new TriggerRegistry();
    gate.register({
      id: 'veto',
      on: 'pre-actuation',
      handle: () => ({ kind: 'block', reason: 'needs approval' }),
    });
    const { client } = fakeClient(['```cmd\nset A1 5\n```', '```cmd\ndone\n```']);
    const session = new AssistSession(bridge, client, { unit, triggers: gate });

    const events = await collect(session.runCommands('write', { approveWrite: () => true }));
    const write = loopEvents(events).find((e) => e.type === 'write-result');
    expect(write).toMatchObject({ result: { ok: false, error: { code: 'blocked' } } });
    expect(bridge.applied).toHaveLength(0); // never actuated
  });

  it('fail-closed: a write with no approver is refused and never actuated', async () => {
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient(['```cmd\nset A1 5\n```', '```cmd\ndone\n```']);
    const session = new AssistSession(bridge, client, { unit });
    // No approveWrite passed → the loop must refuse the write (the DocBridge confirmation contract).
    const events = await collect(session.runCommands('write'));
    const write = loopEvents(events).find((e) => e.type === 'write-result');
    expect(write).toMatchObject({ result: { ok: false, error: { code: 'unapproved' } } });
    expect(bridge.applied).toHaveLength(0);
  });

  it('write-one cap: only maxWritesPerTurn writes actuate in one block', async () => {
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient([
      '```cmd\nset A1 1\nset A2 2\nset A3 3\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });
    const events = await collect(
      session.runCommands('write three', { approveWrite: () => true, maxWritesPerTurn: 2 }),
    );
    expect(bridge.applied).toHaveLength(2); // the third write is capped, not actuated
    expect(loopEvents(events).some((e) => e.type === 'capped')).toBe(true);
  });

  it('suggest → tracked-change with target.matchText (Word)', async () => {
    const bridge = new FakeWordBridge();
    const wordUnit = { connectors: [], surfaceContext: { kind: 'word' as const } };
    const { client } = fakeClient([
      '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is ~99.5% (source needed)."\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit: wordUnit });

    await collect(session.runCommands('flag unsourced claims', { approveWrite: () => true }));
    expect(bridge.applied[0]).toMatchObject({
      kind: 'tracked-change',
      params: {
        target: { matchText: 'The SLA is 99.5%.' },
        text: 'The SLA is ~99.5% (source needed).',
      },
    });
  });

  it('an unknown verb returns a corrective error the next turn self-corrects', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      '```cmd\nsett A1 5\n```', // typo
      '```cmd\nset A1 5\n```', // corrected
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    await collect(session.runCommands('write a cell', { approveWrite: () => true }));

    // The corrective error was fed back on turn 2's query.
    expect(queries[1]).toContain('unknown verb');
    expect(queries[1]).toContain('did you mean');
    // The self-corrected write landed.
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]).toMatchObject({ params: { target: { range: 'A1' } } });
  });

  it('a no-fence turn re-prompts once, then proceeds', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      '**thought** I am still thinking, no commands yet.', // no fence
      '```cmd\nset A1 1\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('write a cell', { approveWrite: () => true }));
    const loop = loopEvents(events);

    expect(loop.some((e) => e.type === 'no-fence')).toBe(true);
    // The re-prompt query nudges for a cmd block.
    expect(queries[1]).toContain('```cmd');
    expect(bridge.applied).toHaveLength(1); // still completed the write
    expect(loop.at(-1)).toMatchObject({ type: 'done' });
  });

  it('stops at maxTurns without done (exhausted)', async () => {
    const bridge = new FakeExcelBridge();
    // Every turn emits a read, never done.
    const { client } = fakeClient(Array(20).fill('```cmd\noutline\n```'));
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('loop forever', { maxTurns: 3 }));
    const loop = loopEvents(events);
    expect(loop.filter((e) => e.type === 'turn-start')).toHaveLength(3);
    expect(loop.at(-1)).toMatchObject({ type: 'exhausted', turns: 3 });
  });

  it('leaves plain ask() unchanged (still streams a grounded answer)', async () => {
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient(['hello world']);
    const session = new AssistSession(bridge, client, { unit });
    const out: SseEvent[] = [];
    for await (const e of session.ask('hi')) out.push(e as SseEvent);
    expect(out.map((e) => e.type)).toContain('token');
  });
});
