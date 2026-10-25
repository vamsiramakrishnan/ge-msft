import { describe, expect, it, vi } from 'vitest';
import { StreamAssistClient } from '@ge/gemini-client';
import {
  asChangeId,
  type ActuationResult,
  type ActuationRequest,
  type DurableCommandProgram,
} from '@ge/contracts';
import { AssistSession, type CommandLoopEvent } from './assist-session.js';
import type { DocBridge } from './bridge.js';
import { MemoryTaskJournalStore, TaskJournal, type TaskJournalStore } from './task-journal.js';
import {
  DurableProgramRunner,
  MemoryDurableProgramStore,
  type DurableProgramStore,
} from './durable-program.js';

const scope = 'sha256:' + 'a'.repeat(64);
const program: DurableCommandProgram = {
  version: 1,
  id: 'month-close',
  surface: 'word',
  stages: [
    {
      id: 'memo',
      dependsOn: [],
      program: 'suggest "Old memo" => "New memo"\nfinish when=verified',
    },
    {
      id: 'summary',
      dependsOn: ['memo'],
      program: 'suggest "Old summary" => "New summary"\nfinish when=verified',
    },
  ],
};
const verified = (id: string): ActuationResult => ({
  ok: true,
  changeId: asChangeId(id),
  kind: 'tracked-change',
  verification: { status: 'verified' },
});
async function collect<T>(events: AsyncGenerator<T>): Promise<T[]> {
  const result: T[] = [];
  for await (const event of events) result.push(event);
  return result;
}
type CheckpointEffect = (request: ActuationRequest) => Promise<void>;
const effect = (id: string): ActuationRequest => ({
  changeId: asChangeId(id),
  kind: 'tracked-change',
  surface: 'word',
  params: { target: { matchText: 'old' }, text: 'new' },
});
async function* complete(
  stage: { id: string },
  checkpointEffect: CheckpointEffect,
): AsyncGenerator<never, ActuationResult[]> {
  yield* [];
  await checkpointEffect(effect(stage.id));
  return [verified(stage.id)];
}

describe('durable command DAG continuation', () => {
  it('reopens after a clean refusal, skips verified stages, and requires a fresh approval for the remainder', async () => {
    const store = new MemoryDurableProgramStore();
    const first = new DurableProgramRunner(store, scope);
    const execute = vi.fn(async function* (stage, checkpointEffect) {
      yield* [];
      await checkpointEffect(effect(stage.id));
      return [
        stage.id === 'memo'
          ? verified('memo')
          : {
              ok: false,
              changeId: asChangeId('summary'),
              kind: 'tracked-change' as const,
              error: { code: 'plan_unapproved', message: 'No' },
            },
      ];
    });
    await collect(first.run(program, execute));
    expect((await first.inspect(program.id))?.stages.map((stage) => stage.state)).toEqual([
      'verified',
      'pending',
    ]);
    const reopened = new DurableProgramRunner(store, scope);
    const resumed = vi.fn(complete);
    await expect(collect(reopened.run(program, resumed))).rejects.toThrow('fresh source');
    expect(resumed).not.toHaveBeenCalled();
    await collect(
      reopened.run(program, resumed, {
        revalidate: async (_, checkpoints) => checkpoints[0]?.receipt?.changeId === 'memo',
      }),
    );
    expect(resumed).toHaveBeenCalledOnce();
    expect(resumed.mock.calls[0]![0].id).toBe('summary');
    expect(
      (await reopened.inspect(program.id))?.stages.every((stage) => stage.state === 'verified'),
    ).toBe(true);
  });
  it('refuses stale completed outputs, modified programs, changed principal/document scopes and foreign receipts', async () => {
    const store = new MemoryDurableProgramStore();
    const runner = new DurableProgramRunner(store, scope);
    await collect(runner.run(program, complete));
    const execute = vi.fn(complete);
    await expect(
      collect(runner.run(program, execute, { revalidate: async () => false })),
    ).rejects.toThrow('fresh source');
    await expect(
      collect(
        runner.run(
          { ...program, stages: [{ ...program.stages[0]!, program: 'suggest "a" => "b"' }] },
          execute,
        ),
      ),
    ).rejects.toThrow('changed');
    const foreign = new DurableProgramRunner(store, 'sha256:' + 'b'.repeat(64));
    expect(await foreign.inspect(program.id)).toBeUndefined();
    expect(execute).not.toHaveBeenCalled();
    await store
      .update(scope, (raw) => {
        const records = structuredClone(raw) as Awaited<ReturnType<typeof runner.inspect>>[];
        records[0]!.stages[0]!.receipt = {
          ...verified('wrong'),
          verification: { status: 'unknown' },
        };
        return records as never;
      })
      .catch(() => undefined);
    expect((await runner.inspect(program.id))?.stages[0]?.receipt?.changeId).toBe('memo');
  });
  it('serializes competing panes and preserves a crashed claim without lease-based replay', async () => {
    const store = new MemoryDurableProgramStore();
    const first = new DurableProgramRunner(store, scope);
    let release!: () => void;
    const barrier = new Promise<void>((resolve) => {
      release = resolve;
    });
    const execute = vi.fn(async function* (_stage, checkpointEffect) {
      yield* [];
      await checkpointEffect(effect('first'));
      await barrier;
      throw new Error('crash');
      return [];
    });
    const pending = collect(first.run(program, execute));
    await vi.waitFor(() => expect(execute).toHaveBeenCalledOnce());
    const duplicate = vi.fn(complete);
    await expect(
      collect(new DurableProgramRunner(store, scope).run(program, duplicate)),
    ).rejects.toThrow('owns this program');
    release();
    await expect(pending).rejects.toThrow('crash');
    await expect(
      collect(first.run(program, duplicate, { revalidate: async () => true })),
    ).rejects.toThrow('owns this program');
    expect(duplicate).not.toHaveBeenCalled();
    expect((await first.inspect(program.id))?.stages[0]?.state).toBe('running');
  });
  it.each(['uncertain', 'unverified', 'checkpoint-pending'] as const)(
    'never replays a %s outcome or advances its dependents',
    async (outcome) => {
      const runner = new DurableProgramRunner(new MemoryDurableProgramStore(), scope);
      const execute = vi.fn(async function* (_stage, checkpointEffect) {
        yield* [];
        await checkpointEffect(effect('first'));
        return [
          {
            ...verified('first'),
            ...(outcome === 'uncertain'
              ? { recoveryPending: true }
              : outcome === 'unverified'
                ? { verification: undefined }
                : { recoveryPending: true }),
          },
        ];
      });
      await collect(runner.run(program, execute));
      expect(execute).toHaveBeenCalledOnce();
      expect((await runner.inspect(program.id))?.stages.map((stage) => stage.state)).toEqual([
        'blocked',
        'pending',
      ]);
      await expect(
        collect(runner.run(program, complete, { revalidate: async () => true })),
      ).rejects.toThrow('cannot be replayed');
    },
  );
  it('fails before executing on persistence failure and leaves an in-flight stage if receipt checkpointing fails', async () => {
    const memory = new MemoryDurableProgramStore();
    let writes = 0;
    const store: DurableProgramStore = {
      durable: true,
      read: (key) => memory.read(key),
      update: async (key, change) => {
        if (++writes >= 4) throw new Error('disk full');
        await memory.update(key, change);
      },
    };
    const execute = vi.fn(complete);
    await expect(
      collect(
        new DurableProgramRunner(
          {
            ...store,
            update: async () => {
              throw new Error('offline');
            },
          },
          scope,
        ).run(program, execute),
      ),
    ).rejects.toThrow('offline');
    expect(execute).not.toHaveBeenCalled();
    await expect(
      collect(new DurableProgramRunner(store, scope).run(program, execute)),
    ).rejects.toThrow('disk full');
    expect(execute).toHaveBeenCalledOnce();
    expect(
      (await new DurableProgramRunner(memory, scope).inspect(program.id))?.stages[0]?.state,
    ).toBe('running');
  });
  it('rejects cycles, duplicate IDs, macros, invalid commands, estate shares and oversized Unicode data', async () => {
    const runner = new DurableProgramRunner(new MemoryDurableProgramStore(), scope);
    const execute = vi.fn(complete);
    for (const stages of [
      [{ ...program.stages[0]!, dependsOn: ['memo'] }],
      [program.stages[0]!, program.stages[0]!],
      [{ ...program.stages[0]!, program: 'def x():\n suggest "a" => "b"\nend' }],
      [{ ...program.stages[0]!, program: 'nonexistent thing' }],
      [{ ...program.stages[0]!, program: 'share report answer' }],
      Array.from({ length: 5 }, (_, i) => ({
        id: `step-${i}`,
        dependsOn: [],
        program: `suggest "a" => "${'😀'.repeat(20000)}"`,
      })),
    ])
      await expect(collect(runner.run({ ...program, stages }, execute))).rejects.toThrow();
    expect(execute).not.toHaveBeenCalled();
  });
  it('does not persist callbacks, approvals, receipt inverse bodies, error text or locations', async () => {
    const store = new MemoryDurableProgramStore();
    const runner = new DurableProgramRunner(store, scope);
    await collect(
      runner.run(
        { ...program, stages: [program.stages[0]!] },
        async function* (_stage, checkpointEffect) {
          yield* [];
          await checkpointEffect(effect('memo'));
          return [
            {
              ...verified('memo'),
              inverse: {
                op: 'restore-text',
                anchor: 'private anchor',
                priorText: 'private inverse',
              } as never,
              location: 'private location',
              verification: { status: 'verified', message: 'private body' },
            },
          ];
        },
      ),
    );
    const saved = JSON.stringify(await store.read(scope));
    expect(saved).not.toMatch(/private|revalidate|approve|inverse|location/);
    expect(saved).toContain('New memo');
  });
  it('cancellation before claim neither dispatches nor creates a saved task', async () => {
    const store = new MemoryDurableProgramStore();
    const runner = new DurableProgramRunner(store, scope);
    const abort = new AbortController();
    abort.abort();
    await expect(
      collect(runner.run(program, complete, { signal: abort.signal })),
    ).rejects.toThrow();
    expect(await runner.inspect(program.id)).toBeUndefined();
  });
});

describe('explicit durable stage reconciliation', () => {
  it('keeps ambiguity blocked, guards in-flight work, then permits inspected not-applied work with a new attempt', async () => {
    const runner = new DurableProgramRunner(new MemoryDurableProgramStore(), scope);
    await collect(
      runner.run(program, async function* (_stage, checkpointEffect) {
        yield* [];
        await checkpointEffect(effect('uncertain'));
        return [
          {
            ok: false,
            kind: 'tracked-change',
            changeId: asChangeId('uncertain'),
            recoveryPending: true,
          },
        ];
      }),
    );
    const inspect = vi.fn(async () => ({ decision: 'not-applied' as const }));
    await expect(
      runner.reconcile(program.id, 'memo', inspect, { confirmStopped: true }),
    ).rejects.toThrow('five minutes');
    expect(inspect).not.toHaveBeenCalled();
    vi.useFakeTimers();
    try {
      vi.setSystemTime(Date.now() + 6 * 60 * 1000);
      await runner.reconcile(program.id, 'memo', async () => ({ decision: 'unresolved' }), {
        confirmStopped: true,
      });
      expect((await runner.inspect(program.id))?.stages[0]?.state).toBe('blocked');
      await expect(collect(runner.run(program, complete))).rejects.toThrow('cannot be replayed');
      await runner.reconcile(program.id, 'memo', inspect, { confirmStopped: true });
      expect((await runner.inspect(program.id))?.stages[0]).toMatchObject({
        state: 'pending',
        inspections: [
          { decision: 'unresolved' },
          { decision: 'not-applied', changeId: 'uncertain' },
        ],
      });
      await collect(runner.run(program, complete));
      expect(
        (await runner.inspect(program.id))?.stages.every((stage) => stage.state === 'verified'),
      ).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
  it('requires a correlated verified native receipt before satisfying an applied stage', async () => {
    const runner = new DurableProgramRunner(new MemoryDurableProgramStore(), scope);
    await collect(
      runner.run(program, async function* (_stage, checkpointEffect) {
        yield* [];
        await checkpointEffect(effect('applied'));
        return [{ ok: true, kind: 'tracked-change', changeId: asChangeId('applied') }];
      }),
    );
    vi.useFakeTimers();
    try {
      vi.setSystemTime(Date.now() + 6 * 60 * 1000);
      await expect(
        runner.reconcile(program.id, 'memo', async () => ({ decision: 'not-applied' }), {
          confirmStopped: true,
        }),
      ).rejects.toThrow('known applied');
      await expect(
        runner.reconcile(
          program.id,
          'memo',
          async () => ({ decision: 'verified', receipt: verified('foreign') }),
          { confirmStopped: true },
        ),
      ).rejects.toThrow('correlated');
      await runner.reconcile(
        program.id,
        'memo',
        async () => ({ decision: 'verified', receipt: verified('applied') }),
        { confirmStopped: true },
      );
      expect((await runner.inspect(program.id))?.stages[0]?.state).toBe('verified');
      const execute = vi.fn(complete);
      await expect(collect(runner.run(program, execute))).rejects.toThrow('fresh source');
      await collect(runner.run(program, execute, { revalidate: async () => true }));
      expect(execute).toHaveBeenCalledOnce();
    } finally {
      vi.useRealTimers();
    }
  });
  it('does not classify a foreign receipt as this stage success', async () => {
    const runner = new DurableProgramRunner(new MemoryDurableProgramStore(), scope);
    await collect(
      runner.run(program, async function* (_stage, checkpointEffect) {
        yield* [];
        await checkpointEffect(effect('expected'));
        return [verified('foreign')];
      }),
    );
    expect((await runner.inspect(program.id))?.stages[0]?.state).toBe('blocked');
  });
});

function fixture(
  store: DurableProgramStore = new MemoryDurableProgramStore(),
  validatePrincipal?: () => Promise<void>,
  journalStore: TaskJournalStore = new MemoryTaskJournalStore(),
) {
  let documentKey = 'document-1';
  const actuate = vi.fn(async (request) => verified(request.changeId));
  const bridge: DocBridge = {
    surface: 'word',
    documentKey: () => documentKey,
    getCapabilities: () => ({
      surface: 'word',
      contextKinds: [],
      actuations: [{ kind: 'tracked-change', surface: 'word', title: 'Tracked', reversible: true }],
    }),
    listContext: async () => [],
    resolveContext: async () => [],
    actuate,
  };
  const client = new StreamAssistClient(
    { getAccessToken: async () => 'unused' },
    { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'test' },
    vi.fn() as never,
  );
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'word' } },
    durablePrograms: store,
    ...(validatePrincipal ? { validatePrincipal } : {}),
    taskJournal: new TaskJournal(journalStore, scope),
    autoAttach: [],
  });
  return {
    session,
    actuate,
    changeDocument: () => {
      documentKey = 'document-2';
    },
  };
}
describe('durable DAG uses the real shared session executor', () => {
  it('preserves approval, executes once, and reopens checkpoints without dispatching their verified stages', async () => {
    const store = new MemoryDurableProgramStore();
    const { session, actuate } = fixture(store);
    const approve = vi.fn(async () => true);
    const first = await collect(session.runDurableProgram(program, { approvePlan: approve }));
    expect(actuate).toHaveBeenCalledTimes(2);
    expect(approve).toHaveBeenCalledTimes(2);
    expect(first.some((event) => event.type === 'done')).toBe(true);
    const reopened = fixture(store);
    await collect(
      reopened.session.runDurableProgram(program, {
        approvePlan: approve,
        revalidate: async () => true,
      }),
    );
    expect(reopened.actuate).not.toHaveBeenCalled();
    expect(approve).toHaveBeenCalledTimes(2);
  });
  it('requires fresh approval after a rejected stage and refuses changes to the active document', async () => {
    const { session, actuate, changeDocument } = fixture();
    await collect(session.runDurableProgram(program, { approvePlan: () => false }));
    expect(actuate).not.toHaveBeenCalled();
    const events = await collect(
      session.runDurableProgram(program, {
        approvePlan: () => {
          changeDocument();
          return true;
        },
      }),
    );
    expect(actuate).not.toHaveBeenCalled();
    expect(
      events.some(
        (event) => event.type === 'write-result' && event.result.error?.code === 'target_conflict',
      ),
    ).toBe(true);
  });
  it('refuses multiple mutations in a stage before dispatch', async () => {
    const { session, actuate } = fixture();
    const oneStage = {
      ...program,
      stages: [
        {
          ...program.stages[0]!,
          program: 'suggest "a" => "b"\nsuggest "c" => "d"\nfinish when=verified',
        },
      ],
    };
    const approve = vi.fn(async () => true);
    await collect(session.runDurableProgram(oneStage, { approvePlan: approve }));
    expect(actuate).not.toHaveBeenCalled();
    expect(approve).not.toHaveBeenCalled();
  });
  it('does not turn an invalid program prefix into a verified checkpoint', async () => {
    const { session, actuate } = fixture();
    const events = await collect(
      session.runDurableProgram(
        {
          ...program,
          stages: [
            {
              ...program.stages[0]!,
              program: 'suggest "a" => "b"\nread Missing\nfinish when=verified',
            },
          ],
        },
        { approvePlan: () => true },
      ),
    );
    expect(actuate).not.toHaveBeenCalled();
    expect(
      events
        .filter(
          (event): event is Extract<CommandLoopEvent, { type: 'durable-progress' }> =>
            event.type === 'durable-progress',
        )
        .at(-1)?.state,
    ).toBe('blocked');
  });
  it('unblocks the independent effect journal only after an explicit not-applied inspection', async () => {
    const { session, actuate } = fixture();
    actuate.mockImplementationOnce(async () => {
      throw new Error('native interrupted');
    });
    await collect(session.runDurableProgram(program, { approvePlan: () => true }));
    expect((await session.inspectDurableProgram(program.id))?.stages[0]?.state).toBe('blocked');
    vi.useFakeTimers();
    try {
      vi.setSystemTime(Date.now() + 6 * 60 * 1000);
      await session.reconcileDurableStage(
        program.id,
        'memo',
        async () => ({ decision: 'not-applied' }),
        { confirmStopped: true },
      );
      const approve = vi.fn(async () => true);
      await collect(session.runDurableProgram(program, { approvePlan: approve }));
      expect(approve).toHaveBeenCalledTimes(2);
      expect(actuate).toHaveBeenCalledTimes(3);
      expect(
        (await session.inspectDurableProgram(program.id))?.stages.every(
          (stage) => stage.state === 'verified',
        ),
      ).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
  it('refuses run, inspect and reconcile after an account switch', async () => {
    let samePrincipal = true;
    const validatePrincipal = async () => {
      if (!samePrincipal) throw new Error('account switched');
    };
    const { session, actuate } = fixture(new MemoryDurableProgramStore(), validatePrincipal);
    actuate.mockImplementationOnce(async () => {
      throw new Error('interrupted');
    });
    await collect(session.runDurableProgram(program, { approvePlan: () => true }));
    samePrincipal = false;
    await expect(
      collect(session.runDurableProgram(program, { approvePlan: () => true })),
    ).rejects.toThrow('account switched');
    await expect(session.inspectDurableProgram(program.id)).rejects.toThrow('account switched');
    const inspect = vi.fn(async () => ({ decision: 'not-applied' as const }));
    await expect(
      session.reconcileDurableStage(program.id, 'memo', inspect, { confirmStopped: true }),
    ).rejects.toThrow('account switched');
    expect(inspect).not.toHaveBeenCalled();
    expect(actuate).toHaveBeenCalledOnce();
  });
  it('rechecks account identity after durable effect intent persistence and before native dispatch', async () => {
    const effects = new MemoryTaskJournalStore();
    let samePrincipal = true;
    let writes = 0;
    const journalStore: TaskJournalStore = {
      durable: true,
      read: (scope) => effects.read(scope),
      update: async (scope, change) => {
        await effects.update(scope, change);
        if (++writes === 1) samePrincipal = false;
      },
    };
    const { session, actuate } = fixture(
      new MemoryDurableProgramStore(),
      async () => {
        if (!samePrincipal) throw new Error('account switched');
      },
      journalStore,
    );
    const events = await collect(session.runDurableProgram(program, { approvePlan: () => true }));
    expect(writes).toBeGreaterThanOrEqual(1);
    expect(actuate).not.toHaveBeenCalled();
    expect(
      events.some(
        (event) => event.type === 'write-result' && event.result.error?.code === 'identity_changed',
      ),
    ).toBe(true);
  });
});
