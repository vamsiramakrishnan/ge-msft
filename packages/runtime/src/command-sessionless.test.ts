import { describe, expect, it, vi } from 'vitest';
import {
  asSessionId,
  gridForRequest,
  makeCellSnapshot,
  type ActuationRequest,
  type AssistRequest,
  type CellValue,
  type SseEvent,
} from '@ge/contracts';
import type { ResolvedGrounding, StreamAssistClient, StreamOptions } from '@ge/gemini-client';
import {
  AssistSession,
  type AssistSessionOptions,
  type CommandLoopEvent,
} from './assist-session.js';
import type { DocBridge } from './bridge.js';

type Event = SseEvent | CommandLoopEvent;
type Response = string | { text: string; sessionId: string };
type Invocation = { request: AssistRequest; options: StreamOptions };
const command = (text: string): string => `\`\`\`cmd\n${text}\n\`\`\``;
const planner = '```plan\nintent review\nsurface excel\nstep Inspect\n```';
const completed = (events: Event[]): boolean =>
  events.some((event) => event.type === 'done' && 'turn' in event);
async function collect(generator: AsyncGenerator<Event>): Promise<Event[]> {
  const events: Event[] = [];
  for await (const event of generator) events.push(event);
  return events;
}

/** Runtime integration: model responses and Office are simulated; no provider/network call occurs. */
function fixture(
  responses: Response[],
  options: Partial<AssistSessionOptions> = {},
  mutateStreamOptions?: (options: StreamOptions) => void,
) {
  const calls: Invocation[] = [];
  const landed: ActuationRequest[] = [];
  let cells: CellValue[][] = [[0]];
  let captures = 0;
  const bridge: DocBridge = {
    surface: 'excel',
    getCapabilities: () => ({
      surface: 'excel',
      contextKinds: ['range', 'sheet'],
      reads: ['read', 'outline'],
      actuations: [
        { kind: 'write-cells', title: 'Write cells', surface: 'excel', reversible: true },
      ],
    }),
    listContext: async () => [],
    resolveContext: async () => [],
    captureDocState: async () => ({
      surface: 'excel',
      version: ++captures,
      capturedAt: new Date(Date.UTC(2026, 8, 7, 0, 0, captures)).toISOString(),
      title: 'Current workbook metadata',
      outline: [],
      inventory: [{ kind: 'sheet', id: 'data', title: 'Data', summary: 'Approved totals' }],
    }),
    readRange: async (selector) => {
      if (selector !== 'Data!A1') throw new Error('Missing range; use Data!A1');
      return [
        {
          ref: {
            id: 'xl:Data!A1',
            kind: 'range',
            surface: 'excel',
            title: 'Approved total',
            live: false,
          },
          value: { as: 'text', text: 'Approved total: 42' },
        },
      ];
    },
    captureCells: (locator) =>
      makeCellSnapshot({
        surface: 'excel',
        documentId: 'test-book',
        objectId: 'results',
        locator,
        values: cells,
      }),
    actuate: async (request) => {
      const before = await bridge.captureCells!(request.params.target!.range!);
      if (request.preconditions?.some((expected) => expected.hash !== before.hash))
        return {
          ok: false,
          kind: request.kind,
          changeId: request.changeId,
          error: { code: 'stale_target', message: 'Destination changed' },
        };
      landed.push(request);
      cells = structuredClone(gridForRequest(request));
      const after = await bridge.captureCells!(request.params.target!.range!);
      return {
        ok: true,
        kind: request.kind,
        changeId: request.changeId,
        verification: { status: 'verified', beforeHash: before.hash, afterHash: after.hash },
      };
    },
  };
  const client = {
    async *stream(request: AssistRequest, streamOptions: StreamOptions): AsyncGenerator<SseEvent> {
      calls.push({ request: structuredClone(request), options: structuredClone(streamOptions) });
      const response = responses[calls.length - 1];
      if (response === undefined) throw new Error('Unexpected model inference');
      mutateStreamOptions?.(streamOptions);
      yield { type: 'token', text: typeof response === 'string' ? response : response.text };
      // A buggy/custom transport could return a session id. The runtime must not adopt or
      // attach it to sessionless machine output, even though the real client also strips it.
      yield {
        type: 'provenance',
        payload: {
          agentId: 'test-agent',
          identity: 'test-owner',
          timestamp: '2026-09-07T00:00:00Z',
          sources: [],
          contentHash: 'test-content',
          sessionId: asSessionId(
            typeof response === 'string' ? 'rogue-machine-session' : response.sessionId,
          ),
        },
      };
      yield { type: 'done' };
    },
  } as unknown as StreamAssistClient;
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'excel' } },
    resumeSessionId: asSessionId('chat-start'),
    ...options,
  });
  return { session, calls, landed, cells: () => cells, captures: () => captures };
}

describe('sessionless command and planner execution', () => {
  it('keeps sessionless provenance isolated when a custom adapter mutates its stream options', async () => {
    const mutate = vi.fn((options: StreamOptions) => {
      options.isSessionLess = false;
    });
    const f = fixture([command('set Results!A1 42\nfinish when=verified')], {}, mutate);
    const observedSessions: unknown[] = [];
    f.session.hooks.register({
      id: 'capture-provenance-session',
      on: 'model:event',
      mode: 'observe',
      handle({ event }) {
        if (event.type === 'provenance') observedSessions.push(event.payload.sessionId);
      },
    });
    const events = await collect(
      f.session.runCommands('Set the approved total', { approvePlan: () => true }),
    );
    expect(completed(events)).toBe(true);
    expect(mutate).toHaveBeenCalledOnce();
    expect(f.calls[0]!.options.isSessionLess).toBe(true);
    expect(observedSessions).toEqual([undefined]);
    expect(f.landed).toHaveLength(1);
    expect(f.landed[0]!.provenance?.sessionId).toBeUndefined();
    for (const event of events)
      if (event.type === 'provenance') expect(event.payload.sessionId).toBeUndefined();
    expect(f.session.sessionId).toBe('chat-start');
    f.session.dispose();
  });

  it('sends self-contained correction turns with the task, protocol, earlier programs/results, live context and grounding', async () => {
    const f = fixture([
      command('read Missing!A1'),
      command('read Data!A1'),
      command('set Results!A1 42\nfinish when=verified'),
    ]);
    const task = 'Set Results!A1 to the approved total; preserve all other cells.';
    f.session.context.add({
      ref: {
        id: 'ctx:approval-policy',
        kind: 'brief',
        surface: 'excel',
        title: 'Relevant policy',
        live: false,
      },
      value: { as: 'text', text: 'Only publish the approved total.' },
    });
    const grounding: ResolvedGrounding = {
      queryParts: [{ text: 'Selected source: approved finance policy' }],
      dataStoreSpecs: [
        {
          dataStore:
            'projects/p/locations/global/collections/default_collection/dataStores/finance',
          filter: 'approved=true',
        },
      ],
    };
    const approvePlan = vi.fn(() => true);
    const events = await collect(f.session.runCommands(task, { grounding, approvePlan }));
    expect(completed(events)).toBe(true);
    expect(f.calls).toHaveLength(3);
    expect(approvePlan).toHaveBeenCalledOnce();
    expect(f.landed).toHaveLength(1);
    expect(f.cells()).toEqual([['42']]);
    for (const call of f.calls) {
      expect(call.options.isSessionLess).toBe(true);
      expect(call.options.session).toBeUndefined();
      expect(call.options.grounding).toEqual(grounding);
      expect(call.options.context).toContainEqual(
        expect.objectContaining({ ref: expect.objectContaining({ id: 'ctx:approval-policy' }) }),
      );
      expect(call.request.query).toContain(task);
      expect(call.request.query).toContain('finish when=verified');
      expect(call.request.query).toContain('Current workbook metadata');
      expect(call.request.query).not.toContain('unchanged="true"');
    }
    expect(f.calls[1]!.request.query).toContain('read Missing!A1');
    expect(f.calls[1]!.request.query).toContain('Missing range; use Data!A1');
    expect(f.calls[2]!.request.query).toContain('read Missing!A1');
    expect(f.calls[2]!.request.query).toContain('read Data!A1');
    expect(f.calls[2]!.request.query).toContain('Approved total: 42');
    expect(f.landed[0]!.provenance).toMatchObject({
      agentId: 'test-agent',
      identity: 'test-owner',
    });
    expect(f.landed[0]!.provenance?.sessionId).toBeUndefined();
    for (const event of events)
      if (event.type === 'provenance') expect(event.payload.sessionId).toBeUndefined();
    expect(f.session.sessionId).toBe('chat-start');
    expect(f.session.executions.list().at(-1)).toMatchObject({
      status: 'completed',
      modelTurns: 3,
    });
    f.session.dispose();
  });

  it('keeps ordinary chat in its conversation while commands and planning never change that session', async () => {
    const f = fixture([
      { text: 'Human conversation answer', sessionId: 'chat-established' },
      { text: command('finish when=verified'), sessionId: 'rogue-command-session' },
      { text: planner, sessionId: 'rogue-planner-session' },
      { text: 'Continuing the human conversation', sessionId: 'chat-established' },
    ]);
    await collect(f.session.ask('Explain this workbook'));
    expect(f.calls[0]!.options.session).toBe('chat-start');
    expect(f.calls[0]!.options.isSessionLess).not.toBe(true);
    expect(f.session.sessionId).toBe('chat-established');
    await collect(f.session.runCommands('Inspect this workbook'));
    expect(f.calls[1]!.options.isSessionLess).toBe(true);
    expect(f.calls[1]!.options.session).toBeUndefined();
    expect(f.session.sessionId).toBe('chat-established');
    await f.session.plan('Review the sheet');
    expect(f.calls[2]!.options.isSessionLess).toBe(true);
    expect(f.calls[2]!.options.session).toBeUndefined();
    expect(f.session.sessionId).toBe('chat-established');
    await collect(f.session.ask('Continue that explanation'));
    expect(f.calls[3]!.options.session).toBe('chat-established');
    expect(f.calls[3]!.options.isSessionLess).not.toBe(true);
    f.session.dispose();
  });

  it('starts a fresh command capsule for every task without leaking prior instructions or receipts', async () => {
    const f = fixture([
      command('read Data!A1'),
      command('finish when=verified'),
      command('finish when=verified'),
    ]);
    await collect(f.session.runCommands('Task alpha: inspect confidential total ZULU-123'));
    await collect(f.session.runCommands('Task beta: inspect layout only'));
    expect(f.calls).toHaveLength(3);
    expect(f.calls[1]!.request.query).toContain('ZULU-123');
    expect(f.calls[1]!.request.query).toContain('Approved total: 42');
    expect(f.calls[2]!.request.query).toContain('Task beta: inspect layout only');
    expect(f.calls[2]!.request.query).not.toContain('ZULU-123');
    expect(f.calls[2]!.request.query).not.toContain('Approved total: 42');
    expect(f.session.sessionId).toBe('chat-start');
    f.session.dispose();
  });

  it('rejects command capsule overflow before spending a model call or mutating the document', async () => {
    const f = fixture([], { commandCapsuleBytes: 1024 });
    const approvePlan = vi.fn(() => true);
    await expect(
      collect(
        f.session.runCommands('Preserve these required constraints. '.repeat(200), { approvePlan }),
      ),
    ).rejects.toThrow(/capsule|budget|context/i);
    expect(f.calls).toHaveLength(0);
    expect(approvePlan).not.toHaveBeenCalled();
    expect(f.landed).toHaveLength(0);
    expect(f.session.executions.list().at(-1)).toMatchObject({ status: 'failed', modelTurns: 0 });
    f.session.dispose();
  });

  it('rejects session-bound uploaded grounding files before a command model call', async () => {
    const f = fixture([]);
    const events = await collect(
      f.session.runCommands('Read the uploaded file', {
        grounding: { fileIds: ['session-upload-1'] },
      }),
    );
    expect(completed(events)).toBe(false);
    expect(events).toContainEqual(
      expect.objectContaining({ type: 'error', code: 'invalid_request' }),
    );
    expect(f.calls).toHaveLength(0);
    expect(f.landed).toHaveLength(0);
    expect(f.session.executions.list().at(-1)).toMatchObject({ status: 'failed', modelTurns: 0 });
    f.session.dispose();
  });

  it('rejects session-bound uploaded grounding files before a planner model call', async () => {
    const f = fixture([]);
    await f.session.plan('Review the uploaded file', {
      grounding: { fileIds: ['session-upload-1'] },
    });
    expect(f.calls).toHaveLength(0);
    expect(f.session.executions.list().at(-1)).toMatchObject({ status: 'failed', modelTurns: 0 });
    f.session.dispose();
  });

  it('retains explicit conversation-mode compatibility including uploaded grounding', async () => {
    const f = fixture([command('finish when=verified')], { commandSessionMode: 'conversation' });
    const grounding = { fileIds: ['session-upload-1'] };
    const events = await collect(f.session.runCommands('Inspect uploaded context', { grounding }));
    expect(completed(events)).toBe(true);
    expect(f.calls).toHaveLength(1);
    expect(f.calls[0]!.options.isSessionLess).not.toBe(true);
    expect(f.calls[0]!.options.session).toBe('chat-start');
    expect(f.calls[0]!.options.grounding).toEqual(grounding);
    expect(f.session.sessionId).toBe('rogue-machine-session');
    f.session.dispose();
  });
});
