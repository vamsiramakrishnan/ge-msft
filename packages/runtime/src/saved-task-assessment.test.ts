import { describe, expect, it, vi } from 'vitest';
import { StreamAssistClient } from '@ge/gemini-client';
import {
  contentHash,
  type ActuationRequest,
  type DurableCommandProgram,
  type TextTargetCapture,
} from '@ge/contracts';
import { AssistSession } from './assist-session.js';
import type { DocBridge } from './bridge.js';
import { MemoryTaskJournalStore, TaskJournal } from './task-journal.js';
import {
  DurableProgramRunner,
  MemoryDurableProgramStore,
  type DurableProgramStore,
} from './durable-program.js';

async function collect<T>(events: AsyncGenerator<T>): Promise<T[]> {
  const values: T[] = [];
  for await (const event of events) values.push(event);
  return values;
}
const program: DurableCommandProgram = {
  version: 1,
  id: 'authored-task',
  surface: 'word',
  evidence: { kind: 'literal' },
  stages: [
    {
      id: 'first',
      dependsOn: [],
      program:
        '/fill-content-control contentControlId=first text="New first"\nfinish when=verified',
    },
    {
      id: 'second',
      dependsOn: ['first'],
      program:
        '/fill-content-control contentControlId=second text="New second"\nfinish when=verified',
    },
  ],
};
async function fixture(store: DurableProgramStore = new MemoryDurableProgramStore()) {
  let key = 'saved-document';
  let principal = true;
  let uncertain = false;
  let readOnly = false;
  const values = new Map([
    ['first', 'Old first'],
    ['second', 'Old second'],
  ]);
  const capture = vi.fn(async (ref): Promise<TextTargetCapture> => {
    if (ref.hostRef?.type !== 'word.contentControl') throw new Error('Wrong target');
    const id = ref.hostRef.contentControlId;
    const text = values.get(id);
    if (text === undefined) throw new Error('Unknown object');
    return {
      text,
      target: {
        surface: 'word',
        documentId: await contentHash({ surface: 'word', documentKey: key }),
        objectId: id,
        locator: `word:cc:${id}`,
        revision: await contentHash({ text }),
        scope: 'word-content-control',
      },
    };
  });
  const actuate = vi.fn(async (request: ActuationRequest) => {
    if (uncertain)
      return { ok: false, kind: request.kind, changeId: request.changeId, recoveryPending: true };
    const id = request.params.target?.contentControlId;
    if (!id || !values.has(id)) throw new Error('Unknown object');
    if (request.params.target?.expectedTextHash !== (await contentHash({ text: values.get(id) })))
      return {
        ok: false,
        kind: request.kind,
        changeId: request.changeId,
        error: { code: 'target_conflict', message: 'Native target changed after preview.' },
      };
    values.set(id, request.params.text!);
    return {
      ok: true,
      kind: request.kind,
      changeId: request.changeId,
      verification: {
        status: 'verified' as const,
        afterHash: await contentHash({ text: request.params.text! }),
      },
    };
  });
  const bridge: DocBridge = {
    surface: 'word',
    documentKey: () => key,
    getCapabilities: () => ({
      surface: 'word',
      contextKinds: [],
      actuations: readOnly
        ? []
        : [
            {
              surface: 'word',
              kind: 'fill-content-control',
              title: 'Fill control',
              reversible: true,
            },
          ],
    }),
    listContext: async () => [],
    resolveContext: async () => [],
    captureTextTarget: capture,
    actuate,
  };
  const client = new StreamAssistClient(
    { getAccessToken: async () => 'unused' },
    { assistant: { project: 'p', engine: 'e', location: 'eu' }, identity: 'test' },
    vi.fn() as never,
  );
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'word' } },
    autoAttach: [],
    durablePrograms: store,
    taskJournal: new TaskJournal(new MemoryTaskJournalStore(), 'sha256:' + 'a'.repeat(64)),
    validatePrincipal: async () => {
      if (!principal) throw new Error('account switched');
    },
  });
  return {
    session,
    store,
    values,
    actuate,
    capture,
    changePrincipal: () => {
      principal = false;
    },
    changeDocument: () => {
      key = 'other-document';
    },
    makeUncertain: () => {
      uncertain = true;
    },
    makeReadOnly: () => {
      readOnly = true;
    },
  };
}

describe('saved authored task operator APIs use native currentness, not canned true validators', () => {
  it('saves and lists metadata without executing, then resumes partial work with fresh native output checks and approval', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    expect(f.actuate).not.toHaveBeenCalled();
    expect(await f.session.listDurablePrograms()).toEqual([
      {
        id: 'authored-task',
        surface: 'word',
        updatedAt: expect.any(String),
        stageCount: 2,
        verifiedStages: 0,
        state: 'ready',
        durable: false,
      },
    ]);
    const summary = JSON.stringify(await f.session.listDurablePrograms());
    expect(summary).not.toMatch(/New first|program|contentControlId/);
    expect((await f.session.assessDurableProgram(program.id))?.canResume).toBe(true);
    let approvals = 0;
    await collect(
      f.session.runInspectedDurableProgram(program.id, { approvePlan: () => ++approvals === 1 }),
    );
    expect(f.actuate).toHaveBeenCalledOnce();
    expect((await f.session.listDurablePrograms())[0]?.state).toBe('partial');
    const approve = vi.fn(async () => true);
    await collect(f.session.runInspectedDurableProgram(program.id, { approvePlan: approve }));
    expect(approve).toHaveBeenCalledOnce();
    expect(f.actuate).toHaveBeenCalledTimes(2);
    expect(f.capture).toHaveBeenCalled();
    expect(
      (await f.session.assessDurableProgram(program.id))?.stages.map((stage) => stage.status),
    ).toEqual(['current', 'current']);
  });
  it('refuses manual output changes before requesting another approval or dispatching', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    let approvals = 0;
    await collect(
      f.session.runInspectedDurableProgram(program.id, { approvePlan: () => ++approvals === 1 }),
    );
    f.values.set('first', 'Manual edit');
    const assessment = await f.session.assessDurableProgram(program.id);
    expect(assessment?.canResume).toBe(false);
    expect(assessment?.stages[0]?.status).toBe('changed');
    const approve = vi.fn(async () => true);
    await expect(
      collect(f.session.runInspectedDurableProgram(program.id, { approvePlan: approve })),
    ).rejects.toThrow('changed');
    expect(approve).not.toHaveBeenCalled();
    expect(f.actuate).toHaveBeenCalledOnce();
  });
  it('checks earlier outputs again after approval, without writing the remaining stage', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    let approvals = 0;
    await collect(
      f.session.runInspectedDurableProgram(program.id, { approvePlan: () => ++approvals === 1 }),
    );
    const events = await collect(
      f.session.runInspectedDurableProgram(program.id, {
        approvePlan: () => {
          f.values.set('first', 'Edit during approval');
          return true;
        },
      }),
    );
    expect(f.actuate).toHaveBeenCalledOnce();
    expect(events).toContainEqual(
      expect.objectContaining({
        type: 'write-result',
        result: expect.objectContaining({
          ok: false,
          error: expect.objectContaining({ code: 'stale_source' }),
        }),
      }),
    );
  });
  it('pins the native target during preview and preserves a user edit made during approval', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    await collect(
      f.session.runInspectedDurableProgram(program.id, {
        approvePlan: (effects) => {
          expect(effects[0]?.request.params.target).toMatchObject({
            expectedDocumentId: expect.stringMatching(/^sha256:/),
            expectedTextHash: expect.stringMatching(/^sha256:/),
          });
          f.values.set('first', 'A human edited this during approval');
          return true;
        },
      }),
    );
    expect(f.values.get('first')).toBe('A human edited this during approval');
    expect(f.actuate).toHaveBeenCalledOnce();
    expect(f.values.get('second')).toBe('Old second');
  });
  it('keeps legacy or unsupported programs inspectable but does not fabricate source lineage', async () => {
    const f = await fixture();
    const legacy = { ...program, evidence: undefined };
    await new DurableProgramRunner(f.store, 'sha256:' + 'a'.repeat(64)).save(legacy);
    expect(await f.session.inspectDurableProgram(program.id)).toBeDefined();
    expect((await f.session.assessDurableProgram(program.id))?.canResume).toBe(false);
    await expect(f.session.saveDurableProgram(legacy)).rejects.toThrow('evidence declaration');
    await expect(
      f.session.saveDurableProgram({
        ...program,
        id: 'opaque',
        stages: [{ id: 'first', dependsOn: [], program: 'suggest "old" => "new"' }],
      }),
    ).rejects.toThrow('native saved-output inspector');
    expect(f.actuate).not.toHaveBeenCalled();
  });
  it('hides only verified completed tasks while retaining the original ID and no-replay authority', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    await expect(f.session.archiveDurableProgram(program.id)).rejects.toThrow('fully verified');
    await collect(f.session.runInspectedDurableProgram(program.id, { approvePlan: () => true }));
    await f.session.archiveDurableProgram(program.id);
    expect(await f.session.listDurablePrograms()).toEqual([]);
    expect((await f.session.inspectDurableProgram(program.id))?.archivedAt).toEqual(
      expect.any(String),
    );
    expect((await f.session.assessDurableProgram(program.id))?.canResume).toBe(false);
    await expect(f.session.saveDurableProgram(program)).rejects.toThrow('hidden');
    await expect(
      collect(
        f.session.runDurableProgram(program, {
          approvePlan: () => true,
          revalidate: async () => true,
        }),
      ),
    ).rejects.toThrow('hidden');
    expect(f.actuate).toHaveBeenCalledTimes(2);
  });
  it('keeps uncertain stages blocked and does not hide or replay them', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    f.makeUncertain();
    await collect(f.session.runInspectedDurableProgram(program.id, { approvePlan: () => true }));
    expect((await f.session.assessDurableProgram(program.id))?.canResume).toBe(false);
    await expect(f.session.archiveDurableProgram(program.id)).rejects.toThrow('fully verified');
    await expect(
      collect(f.session.runInspectedDurableProgram(program.id, { approvePlan: () => true })),
    ).rejects.toThrow();
    expect(f.actuate).toHaveBeenCalledOnce();
  });
  it('retains cleanup failures as blocked receipts rather than claiming a resumable verified task', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    f.actuate.mockImplementationOnce(async (request) => ({
      ok: true,
      kind: request.kind,
      changeId: request.changeId,
      verification: { status: 'verified', afterHash: await contentHash({ text: 'New first' }) },
      trackingModeRestoreFailed: true,
    }));
    await collect(f.session.runInspectedDurableProgram(program.id, { approvePlan: () => true }));
    const record = await f.session.inspectDurableProgram(program.id);
    expect(record?.stages[0]).toMatchObject({
      state: 'blocked',
      receipt: { trackingModeRestoreFailed: true },
    });
    expect((await f.session.assessDurableProgram(program.id))?.canResume).toBe(false);
    await expect(f.session.archiveDurableProgram(program.id)).rejects.toThrow('fully verified');
  });
  it('pins authentication and document identity on every read-only saved task operation', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    f.changePrincipal();
    await expect(f.session.listDurablePrograms()).rejects.toThrow('account switched');
    await expect(f.session.assessDurableProgram(program.id)).rejects.toThrow('account switched');
    await expect(f.session.archiveDurableProgram(program.id)).rejects.toThrow('account switched');
    expect(f.actuate).not.toHaveBeenCalled();
    const g = await fixture();
    await g.session.saveDurableProgram(program);
    g.changeDocument();
    await expect(g.session.listDurablePrograms()).rejects.toThrow('Rebind');
  });
  it('explains a current capability or release-profile refusal before offering resume', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    f.makeReadOnly();
    const assessment = await f.session.assessDurableProgram(program.id);
    expect(assessment?.canResume).toBe(false);
    expect(assessment?.blockers.join(' ')).toContain('current host or release profile');
    const approve = vi.fn(async () => true);
    await expect(
      collect(f.session.runInspectedDurableProgram(program.id, { approvePlan: approve })),
    ).rejects.toThrow('current host or release profile');
    expect(approve).not.toHaveBeenCalled();
    expect(f.actuate).not.toHaveBeenCalled();
  });
  it('bounds a stalled output getter and releases the session inspection owner', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    await collect(f.session.runInspectedDurableProgram(program.id, { approvePlan: () => true }));
    vi.useFakeTimers();
    try {
      f.capture.mockImplementationOnce(() => new Promise<TextTargetCapture>(() => undefined));
      const captured = f.capture.mock.calls.length;
      const pending = f.session.assessDurableProgram(program.id);
      await vi.waitFor(() => expect(f.capture.mock.calls.length).toBeGreaterThan(captured));
      await vi.advanceTimersByTimeAsync(2001);
      const result = await pending;
      expect(result?.canResume).toBe(false);
      expect(result?.stages[0]?.status).toBe('unsupported');
      expect(await f.session.listDurablePrograms()).toHaveLength(1);
      expect((await f.session.assessDurableProgram(program.id))?.canResume).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
  it('cancels a stalled pre-approval target read without retaining pane ownership or dispatching', async () => {
    const f = await fixture();
    await f.session.saveDurableProgram(program);
    const source = f.capture.getMockImplementation()!;
    let release!: () => void;
    f.capture.mockImplementationOnce(
      (ref) =>
        new Promise<TextTargetCapture>((resolve) => {
          release = () => {
            void source(ref).then(resolve);
          };
        }),
    );
    const signal = new AbortController();
    const approve = vi.fn(async () => true);
    const run = collect(
      f.session.runInspectedDurableProgram(program.id, {
        signal: signal.signal,
        approvePlan: approve,
      }),
    );
    const settled = run.catch((error) => error);
    await vi.waitFor(() => expect(f.capture).toHaveBeenCalled());
    signal.abort();
    await settled;
    expect(f.actuate).not.toHaveBeenCalled();
    expect(approve).not.toHaveBeenCalled();
    expect(f.session.executions.list().at(-1)?.status).toBe('cancelled');
    expect(await f.session.listDurablePrograms()).toHaveLength(1);
    release();
    await Promise.resolve();
    expect(f.actuate).not.toHaveBeenCalled();
  });
});
