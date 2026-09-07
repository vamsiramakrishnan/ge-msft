import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { createRequire } from 'node:module';
import { resolve } from 'node:path';
import type * as DuckDB from '@duckdb/duckdb-wasm/blocking';
import {
  asSessionId,
  gridForRequest,
  makeCellSnapshot,
  type ActuationRequest,
  type ActuationResult,
  type AssistRequest,
  type CellValue,
  type DocStateSnapshot,
  type SseEvent,
} from '@ge/contracts';
import { validateQuery, type ComputeEngine } from '@ge/compute';
import type { StreamAssistClient } from '@ge/gemini-client';
import { artifactToIPC, arrowRows, ENGINE_SETTINGS } from '../../compute/src/arrow.js';
import { compileAnalysisProgram, type AnalysisProgram } from './analysis-program.js';
import { AssistSession, type CommandLoopEvent } from './assist-session.js';
import type { DocBridge } from './bridge.js';

// These are acceptance tests of the real runtime and real DuckDB WASM calculations. The model
// response and Office bridge are simulated. Turn/byte counts are not live-provider latency or tokens.
const require = createRequire(import.meta.url);
const duck = require('@duckdb/duckdb-wasm/blocking') as typeof DuckDB;
let database: Awaited<ReturnType<typeof duck.createDuckDB>>;
let connection: ReturnType<typeof database.connect>;
const seeded = new Set<string>();

beforeAll(async () => {
  const dist = resolve(require.resolve('@duckdb/duckdb-wasm'), '..');
  database = await duck.createDuckDB(
    {
      mvp: {
        mainModule: resolve(dist, 'duckdb-mvp.wasm'),
        mainWorker: resolve(dist, 'duckdb-node-mvp.worker.cjs'),
      },
    },
    new duck.VoidLogger(),
    duck.NODE_RUNTIME,
  );
  await database.instantiate();
  database.open({ query: { castBigIntToDouble: false, castDecimalToDouble: false } });
  connection = database.connect();
  connection.query(ENGINE_SETTINGS);
}, 30_000);

afterAll(() => {
  connection?.close();
  database?.reset();
});

function engine(): ComputeEngine {
  return {
    async query(raw, tables, signal) {
      signal?.throwIfAborted();
      const start = performance.now();
      const sql = validateQuery(raw);
      const admitted = new Set(tables.map((table) => table.id));
      for (const id of seeded) {
        if (!admitted.has(id)) {
          connection.query(`DROP TABLE ${id}`);
          seeded.delete(id);
        }
      }
      for (const table of tables) {
        if (!seeded.has(table.id)) {
          connection.insertArrowFromIPCStream(artifactToIPC(table), {
            name: table.id,
            create: true,
          });
          seeded.add(table.id);
        }
      }
      expect(connection.getTableNames(sql).every((name) => admitted.has(name))).toBe(true);
      const result = connection.query(`SELECT * FROM (${sql}) bounded_result LIMIT 5001`);
      return {
        columns: result.schema.fields.map((field) => field.name),
        rows: arrowRows(result, 5000),
        truncated: result.numRows > 5000,
        durationMs: performance.now() - start,
      };
    },
    dispose() {},
  };
}

const program: AnalysisProgram = {
  version: 1,
  steps: [
    { op: 'bind', name: 'invoices', action: { kind: 'capture', range: 'Invoices!A1:C4' } },
    { op: 'bind', name: 'payments', action: { kind: 'capture', range: 'Payments!A1:C5' } },
    {
      op: 'bind',
      name: 'result',
      action: {
        kind: 'reconcile',
        spec: {
          left: '$invoices',
          right: '$payments',
          leftKey: 0,
          rightKey: 0,
          leftAmount: 1,
          rightAmount: 1,
          leftCurrency: 2,
          rightCurrency: 2,
          tolerance: '0.001',
        },
      },
    },
    { op: 'materialize', id: '$result', destination: 'Results!A1' },
  ],
};

type Failure =
  | 'stale-source'
  | 'stale-target'
  | 'mismatch'
  | 'unknown'
  | 'checkpoint-before'
  | 'checkpoint-after'
  | 'cancel'
  | 'cancel-checkpoint'
  | 'reject';

function fixture(failure?: Failure, response?: string) {
  const sources: Record<string, CellValue[][]> = {
    'Invoices!A1:C4': [
      ['id', 'amount', 'currency'],
      ['A', '0.30', 'USD'],
      ['B', '9007199254740993.01', 'USD'],
      ['C', '7.00', 'EUR'],
    ],
    'Payments!A1:C5': [
      ['id', 'amount', 'currency'],
      ['A', '0.10', 'USD'],
      ['A', '0.20', 'USD'],
      ['B', '9007199254740993.00', 'USD'],
      ['C', '7.00', 'USD'],
    ],
  };
  const controller = new AbortController();
  const landed: ActuationRequest[] = [];
  let destination: CellValue[][] | undefined;
  let persisted: unknown = [];
  let saves = 0;
  let destinationChanged = false;
  const bridge: DocBridge = {
    surface: 'excel',
    getCapabilities: () => ({
      surface: 'excel',
      contextKinds: ['range', 'sheet'],
      reads: ['outline', 'read', 'inspect'],
      actuations: [
        { kind: 'write-cells', surface: 'excel', title: 'Write cells', reversible: true },
      ],
    }),
    listContext: async () => [],
    resolveContext: async () => [],
    captureCells: async (locator) => {
      let values = sources[locator];
      if (!values) {
        if (!locator.startsWith('Results!')) throw new Error(`Unknown range ${locator}`);
        const rows = Number(/:(?:[A-Z]+)(\d+)$/.exec(locator)?.[1] ?? 1);
        // The reconciliation contract produces eight output columns.
        values = destination ?? Array.from({ length: rows }, () => Array<CellValue>(8).fill(null));
        if (destinationChanged)
          values = values.map((row, i) => (i === 0 ? ['coauthor edit', ...row.slice(1)] : row));
      }
      return makeCellSnapshot({
        surface: 'excel',
        documentId: 'workbook-efficiency',
        objectId: locator.split('!')[0],
        locator,
        values,
      });
    },
    recoveryStorage: {
      load: async () => structuredClone(persisted),
      save: async (_owner, value) => {
        saves++;
        if (
          (failure === 'checkpoint-before' && saves === 1) ||
          (failure === 'checkpoint-after' && saves === 2)
        )
          throw new Error('Checkpoint unavailable');
        if (failure === 'cancel-checkpoint' && saves === 1) controller.abort();
        persisted = structuredClone(value);
      },
    },
    actuate: async (request): Promise<ActuationResult> => {
      for (const expected of request.preconditions ?? []) {
        const current = await bridge.captureCells!(expected.locator);
        if (current.hash !== expected.hash || current.documentId !== expected.documentId)
          return {
            ok: false,
            kind: request.kind,
            changeId: request.changeId,
            error: {
              code: 'stale_source',
              message: 'Source or destination changed after approval',
            },
          };
      }
      const before = await bridge.captureCells!(request.params.target!.range!);
      destination = structuredClone(gridForRequest(request));
      landed.push(request);
      const after = await bridge.captureCells!(request.params.target!.range!);
      return {
        ok: true,
        kind: request.kind,
        changeId: request.changeId,
        verification: {
          status: failure === 'mismatch' || failure === 'unknown' ? failure : 'verified',
          beforeHash: before.hash,
          afterHash: after.hash,
        },
      };
    },
  };
  const approvePlan = vi.fn(() => {
    if (failure === 'stale-source') sources['Invoices!A1:C4']![1]![1] = '0.31';
    if (failure === 'stale-target') destinationChanged = true;
    if (failure === 'cancel') controller.abort();
    return failure !== 'reject';
  });
  const queries: string[] = [];
  const client = {
    async *stream(request: AssistRequest): AsyncGenerator<SseEvent> {
      queries.push(request.query ?? '');
      if (queries.length > 1)
        throw new Error('A terminal program must not request another inference');
      yield {
        type: 'token',
        text: response ?? `\`\`\`cmd\n${compileAnalysisProgram(program)}\n\`\`\``,
      };
      yield { type: 'done' };
    },
  } as unknown as StreamAssistClient;
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'excel' } },
    compute: async () => engine(),
    recoveryOwner: 'test-owner',
  });
  return { session, queries, approvePlan, controller, landed, values: () => destination, bridge };
}

type Event = SseEvent | CommandLoopEvent;
async function collect(generator: AsyncGenerator<Event>): Promise<Event[]> {
  const events: Event[] = [];
  for await (const event of generator) events.push(event);
  return events;
}
const completed = (events: Event[]) =>
  events.some((event) => event.type === 'done' && 'turn' in event);
const byteLength = (text: string) => new TextEncoder().encode(text).byteLength;
function report(label: string, session: AssistSession): void {
  if (process.env.GE_COMMAND_EFFICIENCY_REPORT === '1') {
    const { modelTurns, toolCalls, metrics } = session.executions.list().at(-1)!;
    console.info(JSON.stringify({ scenario: label, modelTurns, toolCalls, ...metrics }));
  }
}

describe('command efficiency acceptance: actual WASM and simulated Office/model', () => {
  it('accepts uppercase verified completion while preserving one inference and one approval', async () => {
    const text = compileAnalysisProgram(program).replace(
      'finish when=verified',
      'FINISH when=verified',
    );
    const f = fixture(undefined, `\`\`\`cmd\n${text}\n\`\`\``);
    const events = await collect(
      f.session.runCommands('Reconcile invoices and payments', { approvePlan: f.approvePlan }),
    );
    expect(completed(events)).toBe(true);
    expect(f.queries).toHaveLength(1);
    expect(f.approvePlan).toHaveBeenCalledOnce();
    expect(f.landed).toHaveLength(1);
    f.session.dispose();
  });

  it.each([{ maxCommandsPerTurn: 1 }, { maxWritesPerTurn: 0 }])(
    'preflights uppercase FINISH before effects: %j',
    async (limits) => {
      const f = fixture();
      const events = await collect(
        f.session.runCommandProgram('set Results!A1 42\nFINISH when=verified', {
          ...limits,
          approvePlan: f.approvePlan,
        }),
      );
      expect(completed(events)).toBe(false);
      expect(f.approvePlan).not.toHaveBeenCalled();
      expect(f.landed).toHaveLength(0);
      expect(events).toContainEqual(expect.objectContaining({ type: 'capped' }));
      f.session.dispose();
    },
  );

  it.each(['finish', 'FINISH'])(
    'preflights a %s terminal generated by macro argument substitution',
    async (terminal) => {
      const f = fixture();
      await collect(
        f.session.runCommandProgram(
          'def closing(completion):\nset Results!A1 42\n$completion when=verified\nend',
        ),
      );
      const events = await collect(
        f.session.runCommandProgram(`closing ${terminal}`, {
          maxCommandsPerTurn: 2,
          approvePlan: f.approvePlan,
        }),
      );
      expect(completed(events)).toBe(false);
      expect(f.approvePlan).not.toHaveBeenCalled();
      expect(f.landed).toHaveLength(0);
      expect(events).toContainEqual(expect.objectContaining({ type: 'capped' }));
      f.session.dispose();
    },
  );

  it.each([
    '```cmd\nclosing',
    'Narration before the program\n```cmd\nclosing\n```',
    '```cmd\nclosing\n```\nNarration after the program',
    '```cmd\nclosing\n```\n```cmd\nset Results!A1 99\n```',
    'cmd\nclosing',
  ])(
    'requires strict original response framing when a macro hides verified completion: %j',
    async (response) => {
      const f = fixture(undefined, response);
      await collect(
        f.session.runCommandProgram('def closing():\nset Results!A1 42\nfinish when=verified\nend'),
      );
      const events = await collect(
        f.session.runCommands('Run closing', { maxTurns: 1, approvePlan: f.approvePlan }),
      );
      expect(completed(events)).toBe(false);
      expect(f.queries).toHaveLength(1);
      expect(f.approvePlan).not.toHaveBeenCalled();
      expect(f.landed).toHaveLength(0);
      expect(f.session.executions.list().at(-1)?.status).toBe('incomplete');
      f.session.dispose();
    },
  );

  it.each(['done', 'share report.txt = "shared payload"'])(
    'rejects %s before any verified-program effect can land',
    async (command) => {
      const f = fixture();
      const events = await collect(
        f.session.runCommandProgram(`set Results!A1 42\n${command}\nfinish when=verified`, {
          approvePlan: f.approvePlan,
        }),
      );
      expect(completed(events)).toBe(false);
      expect(f.approvePlan).not.toHaveBeenCalled();
      expect(f.landed).toHaveLength(0);
      expect(events).toContainEqual(expect.objectContaining({ type: 'capped' }));
      expect(f.session.executions.list().at(-1)?.status).toBe('incomplete');
      f.session.dispose();
    },
  );

  it.each([
    'closing\nclosing',
    'closing\nfinish when=verified',
    'closing\noutline',
    'closing\nset Results!A1 99',
  ])('requires a unique final terminal after full macro expansion: %j', async (commands) => {
    const f = fixture();
    await collect(
      f.session.runCommandProgram('def closing():\nset Results!A1 42\nfinish when=verified\nend'),
    );
    const events = await collect(
      f.session.runCommandProgram(commands, { approvePlan: f.approvePlan }),
    );
    expect(completed(events)).toBe(false);
    expect(f.approvePlan).not.toHaveBeenCalled();
    expect(f.landed).toHaveLength(0);
    expect(events).toContainEqual(expect.objectContaining({ type: 'capped' }));
    expect(f.session.executions.list().at(-1)?.status).toBe('incomplete');
    f.session.dispose();
  });

  it.each([{ maxCommandsPerTurn: 1 }, { maxWritesPerTurn: 0 }])(
    'rejects an oversized verified model program before any approval or mutation: %j',
    async (limits) => {
      const f = fixture();
      const events = await collect(
        f.session.runCommands('Reconcile invoices and payments into Results!A1', {
          ...limits,
          approvePlan: f.approvePlan,
        }),
      );
      expect(completed(events)).toBe(false);
      expect(f.queries).toHaveLength(1);
      expect(f.approvePlan).not.toHaveBeenCalled();
      expect(f.landed).toHaveLength(0);
      expect(f.session.analysis!.state().artifacts).toHaveLength(0);
      expect(f.session.executions.list().at(-1)?.status).toBe('incomplete');
      expect(events).toContainEqual(expect.objectContaining({ type: 'capped' }));
      f.session.dispose();
    },
  );

  it.each([{ maxCommandsPerTurn: 1 }, { maxWritesPerTurn: 0 }])(
    'enforces the same pre-execution budgets on typed SDK programs: %j',
    async (limits) => {
      const f = fixture();
      const events = await collect(
        f.session.runAnalysisProgram(program, { ...limits, approvePlan: f.approvePlan }),
      );
      expect(completed(events)).toBe(false);
      expect(f.queries).toHaveLength(0);
      expect(f.approvePlan).not.toHaveBeenCalled();
      expect(f.landed).toHaveLength(0);
      expect(f.session.analysis!.state().artifacts).toHaveLength(0);
      expect(f.session.executions.list().at(-1)?.status).toBe('incomplete');
      f.session.dispose();
    },
  );

  it.each([{ maxCommandsPerTurn: 2 }, { maxWritesPerTurn: 0 }])(
    'preflights a terminal hidden in a previously registered macro before executing it: %j',
    async (limits) => {
      const f = fixture();
      const registration = await collect(
        f.session.runCommandProgram(
          'def closing():\nset Results!A1 42\noutline\nfinish when=verified\nend',
        ),
      );
      expect(registration).toContainEqual(
        expect.objectContaining({
          type: 'skill-registered',
          name: 'closing',
          result: expect.objectContaining({ ok: true }),
        }),
      );
      const events = await collect(
        f.session.runCommandProgram('closing', { ...limits, approvePlan: f.approvePlan }),
      );
      expect(completed(events)).toBe(false);
      expect(f.queries).toHaveLength(0);
      expect(f.approvePlan).not.toHaveBeenCalled();
      expect(f.landed).toHaveLength(0);
      expect(f.session.executions.list().at(-1)?.status).toBe('incomplete');
      expect(events).toContainEqual(expect.objectContaining({ type: 'capped' }));
      f.session.dispose();
    },
  );

  it('captures, reconciles, approves, writes and verifies in one model turn', async () => {
    const f = fixture();
    const events = await collect(
      f.session.runCommands('Reconcile invoices and payments into Results!A1', {
        approvePlan: f.approvePlan,
      }),
    );
    expect(f.queries).toHaveLength(1);
    expect(f.approvePlan).toHaveBeenCalledOnce();
    expect(f.landed).toHaveLength(1);
    expect(completed(events)).toBe(true);
    expect(
      f
        .values()
        ?.find((row) => row[0] === 'A')
        ?.slice(2, 6),
    ).toEqual(['0.300000', '0.300000', '0.000000', 'matched']);
    expect(
      f
        .values()
        ?.find((row) => row[0] === 'B')
        ?.slice(2, 6),
    ).toEqual(['9007199254740993.010000', '9007199254740993.000000', '0.010000', 'variance']);
    expect(
      f
        .values()
        ?.filter((row) => row[0] === 'C')
        .map((row) => row[5])
        .sort(),
    ).toEqual(['unallocated', 'unpaid']);
    expect(f.session.executions.list().at(-1)).toMatchObject({
      status: 'completed',
      modelTurns: 1,
      metrics: { queryBytes: byteLength(f.queries[0]!), resultInputBytes: 0, resultOutputBytes: 0 },
    });
    expect(f.session.recovery.list()).toMatchObject([{ state: 'applied', canUndo: true }]);
    report('bound-program', f.session);
    f.session.dispose();
  });

  it('executes the identical typed SDK program with zero model calls and one approval', async () => {
    const f = fixture();
    const events = await collect(
      f.session.runAnalysisProgram(program, { approvePlan: f.approvePlan }),
    );
    expect(f.queries).toHaveLength(0);
    expect(f.approvePlan).toHaveBeenCalledOnce();
    expect(f.landed).toHaveLength(1);
    expect(completed(events)).toBe(true);
    expect(f.values()?.find((row) => row[0] === 'B')?.[4]).toBe('0.010000');
    expect(f.session.executions.list().at(-1)).toMatchObject({
      status: 'completed',
      modelTurns: 0,
      metrics: { queryBytes: 0, resultInputBytes: 0, resultOutputBytes: 0 },
    });
    report('typed-sdk', f.session);
    f.session.dispose();
  });

  it.each([
    ['conversation', 'full'],
    ['sessionless', 'full'],
    ['conversation', 'compact'],
    ['sessionless', 'compact'],
  ] as const)(
    'compares the same four-turn artifact handoff protocol in %s mode with %s disclosure',
    async (commandSessionMode, commandDisclosure) => {
      const f = fixture();
      const queries: string[] = [];
      const client = {
        async *stream(request: AssistRequest): AsyncGenerator<SseEvent> {
          queries.push(request.query ?? '');
          const artifacts = session.analysis!.state().artifacts;
          let commands: string;
          switch (queries.length) {
            case 1:
              commands =
                'analyze {"kind":"capture","range":"Invoices!A1:C4"}\nanalyze {"kind":"capture","range":"Payments!A1:C5"}';
              break;
            case 2: {
              const step = program.steps[2]!;
              if (step.op !== 'bind' || step.action.kind !== 'reconcile')
                throw new Error('Invalid test program');
              commands = `analyze ${JSON.stringify({ kind: 'reconcile', spec: { ...step.action.spec, left: artifacts[0]!.id, right: artifacts[1]!.id } })}`;
              break;
            }
            case 3:
              commands = `analyze ${JSON.stringify({ kind: 'materialize', id: session.analysis!.state().selected, destination: 'Results!A1' })}`;
              break;
            case 4:
              commands = 'done';
              break;
            default:
              throw new Error('The legacy transcript exceeded its expected four model calls');
          }
          yield { type: 'token', text: `\`\`\`cmd\n${commands}\n\`\`\`` };
          yield { type: 'done' };
        },
      } as unknown as StreamAssistClient;
      const session = new AssistSession(f.bridge, client, {
        unit: { connectors: [], surfaceContext: { kind: 'excel' } },
        commandDisclosure,
        commandSessionMode,
        compute: async () => engine(),
        recoveryOwner: 'test-owner',
      });
      const events = await collect(
        session.runCommands('Reconcile invoices and payments into Results!A1', {
          approvePlan: f.approvePlan,
        }),
      );
      expect(completed(events)).toBe(true);
      expect(queries).toHaveLength(4);
      expect(f.approvePlan).toHaveBeenCalledOnce();
      expect(f.landed).toHaveLength(1);
      expect(f.values()?.find((row) => row[0] === 'B')?.[4]).toBe('0.010000');
      expect(session.executions.list().at(-1)).toMatchObject({
        status: 'completed',
        modelTurns: 4,
      });
      expect(session.executions.list().at(-1)!.metrics!.resultInputBytes).toBeGreaterThan(0);
      report(`${commandSessionMode}-${commandDisclosure}-artifact-handoffs`, session);
      session.dispose();
      f.session.dispose();
    },
  );

  it.each<Failure>([
    'stale-source',
    'stale-target',
    'mismatch',
    'unknown',
    'checkpoint-before',
    'checkpoint-after',
    'cancel',
    'cancel-checkpoint',
    'reject',
  ])('does not claim completion, infer again or replay after %s', async (failure) => {
    const f = fixture(failure);
    const events: Event[] = [];
    let thrown: unknown;
    try {
      for await (const event of f.session.runCommands(
        'Reconcile invoices and payments into Results!A1',
        { approvePlan: f.approvePlan, signal: f.controller.signal },
      ))
        events.push(event);
    } catch (error) {
      thrown = error;
    }
    const cancelled = failure === 'cancel' || failure === 'cancel-checkpoint';
    if (cancelled) expect(thrown).toMatchObject({ name: 'AbortError' });
    else expect(thrown).toBeUndefined();
    expect(completed(events)).toBe(false);
    expect(f.queries).toHaveLength(1);
    expect(f.approvePlan).toHaveBeenCalledOnce();
    expect(f.landed).toHaveLength(
      ['mismatch', 'unknown', 'checkpoint-after'].includes(failure) ? 1 : 0,
    );
    expect(f.session.executions.list().at(-1)?.status).toBe(cancelled ? 'cancelled' : 'incomplete');
    if (['mismatch', 'unknown'].includes(failure))
      expect(f.session.recovery.list()).toMatchObject([{ state: 'uncertain', canUndo: false }]);
    if (failure === 'checkpoint-after')
      expect(events).toContainEqual(
        expect.objectContaining({
          type: 'write-result',
          result: expect.objectContaining({ ok: true, recoveryPending: true }),
        }),
      );
    f.session.dispose();
  });

  it('keeps an oversized read outside model context and retrieves a precise projected page', async () => {
    const f = fixture();
    const payload = 'Sensitive full-row detail. '.repeat(4000);
    f.bridge.readRange = async () => [
      {
        ref: { id: 'xl:Large!A1', kind: 'range', surface: 'excel', title: 'Large!A1', live: false },
        value: { as: 'text', text: payload },
      },
    ];
    const snapshot: DocStateSnapshot = {
      surface: 'excel',
      version: 1,
      capturedAt: '2026-09-07T00:00:00Z',
      outline: [],
      inventory: Array.from({ length: 20 }, (_, index) => ({
        kind: 'sheet',
        id: `s${index}`,
        title: `Sheet ${index}`,
        summary: '100 rows × 20 columns',
      })),
    };
    let captures = 0;
    f.bridge.captureDocState = async () => ({
      ...snapshot,
      version: ++captures,
      capturedAt: new Date(Date.UTC(2026, 8, 7, 0, 0, captures)).toISOString(),
    });
    const queries: string[] = [];
    const client = {
      async *stream(request: AssistRequest): AsyncGenerator<SseEvent> {
        const query = request.query ?? '';
        queries.push(query);
        if (queries.length === 1) yield { type: 'token', text: '```cmd\nread Large!A1\n```' };
        else if (queries.length === 2) {
          const match = /```result\n([\s\S]*?)\n```/.exec(query);
          expect(match).not.toBeNull();
          const receipt = (JSON.parse(match![1]!) as Array<{ ref: string }>)[0]!;
          expect(receipt.ref).toMatch(/^result:/);
          yield {
            type: 'token',
            text: `\`\`\`cmd\ninspect ${receipt.ref} path=/0/text offset=0 limit=26\n\`\`\``,
          };
        } else yield { type: 'token', text: '```cmd\nfinish when=verified\n```' };
        yield {
          type: 'provenance',
          payload: {
            agentId: 'test-agent',
            identity: 'test-owner',
            timestamp: '2026-09-07T00:00:00Z',
            sources: [],
            contentHash: 'test',
            sessionId: 'test-session' as never,
          },
        };
        yield { type: 'done' };
      },
    } as unknown as StreamAssistClient;
    const session = new AssistSession(f.bridge, client, {
      unit: { connectors: [], surfaceContext: { kind: 'excel' } },
      commandSessionMode: 'conversation',
    });
    const events = await collect(session.runCommands('Read the first phrase from Large!A1'));
    expect(completed(events)).toBe(true);
    expect(queries).toHaveLength(3);
    expect(queries[1]).not.toContain(payload);
    expect(queries[1]).toContain('unchanged="true"');
    expect(byteLength(queries[1]!)).toBeLessThan(1024);
    expect(queries[2]).toContain(payload.slice(0, 26));
    const record = session.executions.list().at(-1)!;
    expect(record.metrics?.resultInputBytes).toBeGreaterThan(100_000);
    expect(record.metrics?.resultOutputBytes).toBeLessThan(1024);
    expect(record.metrics?.snapshotBytesSaved).toBeGreaterThan(1000);
    expect(record.metrics?.queryBytes).toBe(
      queries.reduce((sum, query) => sum + byteLength(query), 0),
    );
    report('projected-read', session);
    session.dispose();
    f.session.dispose();
  });

  it('discloses changed document structure and never reuses a snapshot from a different conversation', async () => {
    const f = fixture();
    let captures = 0;
    let title = 'Invoices';
    f.bridge.captureDocState = async () => ({
      surface: 'excel',
      version: ++captures,
      capturedAt: new Date(Date.UTC(2026, 8, 7, 0, 0, captures)).toISOString(),
      outline: [],
      inventory: [{ kind: 'sheet', id: 'invoices', title, summary: '3 rows × 3 columns' }],
    });
    const queries: string[] = [];
    const client = {
      async *stream(request: AssistRequest): AsyncGenerator<SseEvent> {
        queries.push(request.query ?? '');
        yield { type: 'token', text: '```plan\nintent review\nsurface excel\nstep Inspect\n```' };
        yield { type: 'done' };
      },
    } as unknown as StreamAssistClient;
    const session = new AssistSession(f.bridge, client, {
      unit: { connectors: [], surfaceContext: { kind: 'excel' } },
      resumeSessionId: asSessionId('conversation-a'),
      commandSessionMode: 'conversation',
    });
    await session.plan('Inspect invoices');
    await session.plan('Inspect invoices');
    expect(queries[0]).toContain('Invoices');
    expect(queries[0]).not.toContain('unchanged="true"');
    expect(queries[1]).toContain('unchanged="true"');
    title = 'Invoices renamed by a coauthor';
    await session.plan('Inspect invoices');
    expect(queries[2]).toContain(title);
    expect(queries[2]).not.toContain('unchanged="true"');
    await session.plan('Inspect invoices');
    expect(queries[3]).toContain('unchanged="true"');
    session.resumeSession('conversation-b');
    await session.plan('Inspect invoices');
    expect(queries[4]).toContain(title);
    expect(queries[4]).not.toContain('unchanged="true"');
    expect(captures).toBe(5);
    session.dispose();
    f.session.dispose();
  });

  it('resends the full snapshot after a planner transport failure before delivery', async () => {
    const f = fixture();
    f.bridge.captureDocState = async () => ({
      surface: 'excel',
      version: 1,
      capturedAt: '2026-09-07T00:00:00Z',
      outline: [],
      inventory: [{ kind: 'sheet', id: 'invoices', title: 'Invoices pending delivery' }],
    });
    const queries: string[] = [];
    const client = {
      async *stream(request: AssistRequest): AsyncGenerator<SseEvent> {
        queries.push(request.query ?? '');
        if (queries.length === 1) throw new Error('Transport failed before delivery');
        yield { type: 'token', text: '```plan\nintent review\nsurface excel\nstep Inspect\n```' };
        yield { type: 'done' };
      },
    } as unknown as StreamAssistClient;
    const session = new AssistSession(f.bridge, client, {
      unit: { connectors: [], surfaceContext: { kind: 'excel' } },
      resumeSessionId: asSessionId('conversation-a'),
      commandSessionMode: 'conversation',
    });
    await expect(session.plan('Inspect invoices')).rejects.toThrow(
      'Transport failed before delivery',
    );
    expect(session.executions.list().at(-1)?.status).toBe('failed');
    await session.plan('Inspect invoices');
    expect(queries).toHaveLength(2);
    expect(queries[1]).toContain('Invoices pending delivery');
    expect(queries[1]).not.toContain('unchanged="true"');
    expect(session.executions.list().at(-1)?.status).toBe('completed');
    session.dispose();
    f.session.dispose();
  });
});
