import { describe, expect, it, vi } from 'vitest';
import { makeCellSnapshot, sourceVersion, type CellValue, type SseEvent } from '@ge/contracts';
import type { StreamAssistClient } from '@ge/gemini-client';
import {
  AssistSession,
  type AssistSessionOptions,
  type CommandLoopEvent,
} from './assist-session.js';
import type { DocBridge } from './bridge.js';
import type { AnalysisAction } from './analysis-workspace.js';

type Event = SseEvent | CommandLoopEvent;
const collect = async (stream: AsyncGenerator<Event>): Promise<Event[]> => {
  const events: Event[] = [];
  for await (const event of stream) events.push(event);
  return events;
};

/** Real session approval/verification pipeline; host and compute are deterministic test adapters. */
async function fixture(
  truncated = false,
  capabilityFilter?: AssistSessionOptions['capabilityFilter'],
  rows: CellValue[][] = [],
) {
  let values: CellValue[][] = [['amount'], ['42']];
  const actuate = vi.fn<Parameters<DocBridge['actuate']>, ReturnType<DocBridge['actuate']>>();
  const bridge: DocBridge = {
    surface: 'excel',
    getCapabilities: () => ({
      surface: 'excel',
      contextKinds: ['range'],
      actuations: [{ kind: 'write-cells', surface: 'excel', title: 'Write', reversible: true }],
    }),
    listContext: async () => [],
    resolveContext: async () => [],
    captureCells: async (locator) =>
      makeCellSnapshot({ surface: 'excel', documentId: 'doc', locator, values }),
    actuate,
  };
  const query = vi.fn(async () => {
    throw new Error('Unexpected compute');
  });
  const session = new AssistSession(bridge, {} as StreamAssistClient, {
    unit: { connectors: [], surfaceContext: { kind: 'excel' } },
    compute: async () => ({ query, dispose() {} }),
    ...(capabilityFilter ? { capabilityFilter } : {}),
  });
  const source = await bridge.captureCells!('Data!A1:A2');
  const artifact = await session.analysis!.artifacts.add({
    title: 'No exceptions',
    labels: ['amount'],
    rows,
    sources: [sourceVersion(source)],
    lineage: { operation: 'query', parents: [] },
    truncated,
  });
  const action: AnalysisAction = {
    kind: 'materialize',
    id: artifact.id,
    destination: 'Results!A1',
    whenNonEmpty: true,
  };
  const approvePlan = vi.fn(() => true);
  return {
    session,
    action,
    approvePlan,
    actuate,
    query,
    edit() {
      values = [['amount'], ['43']];
    },
  };
}

describe.each(['direct', 'program'] as const)('conditional materialization through %s', (route) => {
  const run = (f: Awaited<ReturnType<typeof fixture>>): AsyncGenerator<Event> =>
    route === 'direct'
      ? f.session.runAnalysis(f.action, { approvePlan: f.approvePlan })
      : f.session.runCommandProgram(`analyze ${JSON.stringify(f.action)}\nfinish when=verified`, {
          approvePlan: f.approvePlan,
        });

  it('reports a fresh empty result as skipped, with zero approval, mutation, or compute', async () => {
    const f = await fixture();
    const events = await collect(run(f));
    expect(events).toContainEqual(
      expect.objectContaining({
        type: 'read-result',
        result: expect.objectContaining({ status: 'skipped', reason: 'empty-result', effects: 0 }),
      }),
    );
    expect(events.some((event) => event.type === 'done')).toBe(true);
    expect(events.some((event) => event.type === 'plan-preview')).toBe(false);
    expect(f.approvePlan).not.toHaveBeenCalled();
    expect(f.actuate).not.toHaveBeenCalled();
    expect(f.query).not.toHaveBeenCalled();
    expect(f.session.executions.list().at(-1)).toMatchObject({ status: 'completed', effects: [] });
    f.session.dispose();
  });

  it.each(['stale', 'truncated'] as const)(
    'fails an empty %s artifact instead of claiming a safe skip',
    async (failure) => {
      const f = await fixture(failure === 'truncated');
      if (failure === 'stale') f.edit();
      if (route === 'direct')
        await expect(collect(run(f))).rejects.toThrow(
          failure === 'stale' ? 'Source changed' : 'truncated',
        );
      else {
        const events = await collect(run(f));
        expect(events.some((event) => event.type === 'done')).toBe(false);
        expect(events).toContainEqual(
          expect.objectContaining({ type: 'error', code: 'verification_incomplete' }),
        );
      }
      expect(f.approvePlan).not.toHaveBeenCalled();
      expect(f.actuate).not.toHaveBeenCalled();
      expect(f.query).not.toHaveBeenCalled();
      expect(f.session.executions.list().at(-1)?.status).not.toBe('completed');
      f.session.dispose();
    },
  );

  it('skips fresh empty output on a read-only profile without requiring write capability', async () => {
    const f = await fixture(false, (manifest) => ({ ...manifest, actuations: [] }));
    const events = await collect(run(f));
    expect(events.some((event) => event.type === 'done')).toBe(true);
    expect(f.approvePlan).not.toHaveBeenCalled();
    expect(f.actuate).not.toHaveBeenCalled();
    f.session.dispose();
  });
});

describe('analysis capability disclosure', () => {
  it('advertises preview while hiding filtered writes and still blocks direct execution', async () => {
    const f = await fixture(false, (manifest) => ({ ...manifest, actuations: [] }), [['42']]);
    expect(await f.session.getAnalysisCapabilities()).toEqual({ preview: true, write: false });
    await expect(
      collect(f.session.runAnalysis(f.action, { approvePlan: f.approvePlan })),
    ).rejects.toThrow('Cell writes are disabled');
    expect(f.approvePlan).not.toHaveBeenCalled();
    expect(f.actuate).not.toHaveBeenCalled();
    f.session.dispose();
    expect(await f.session.getAnalysisCapabilities()).toEqual({ preview: false, write: false });
  });

  it('advertises writes only when analysis and the effective write capability both exist', async () => {
    const f = await fixture();
    expect(await f.session.getAnalysisCapabilities()).toEqual({ preview: true, write: true });
    const unavailable = new AssistSession(
      { surface: 'excel' } as DocBridge,
      {} as StreamAssistClient,
      { unit: { connectors: [], surfaceContext: { kind: 'excel' } } },
    );
    expect(await unavailable.getAnalysisCapabilities()).toEqual({ preview: false, write: false });
    unavailable.dispose();
    f.session.dispose();
  });
});
