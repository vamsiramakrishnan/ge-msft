import { describe, expect, it, vi } from 'vitest';
import { ArtifactStore } from '@ge/compute';
import { StreamAssistClient } from '@ge/gemini-client';
import {
  contentHash,
  makeCellSnapshot,
  type ActuationRequest,
  type OfficeHandoffPlan,
  type TextTargetCapture,
} from '@ge/contracts';
import { AssistSession } from './assist-session.js';
import type { DocBridge } from './bridge.js';
import { MemoryTaskJournalStore, TaskJournal } from './task-journal.js';
import { MemoryDurableProgramStore } from './durable-program.js';
import { createOfficeApplicationPort, OfficeWorkflowCoordinator } from './office-handoff.js';

async function collect<T>(events: AsyncGenerator<T>): Promise<T[]> {
  const result: T[] = [];
  for await (const event of events) result.push(event);
  return result;
}
async function fixture() {
  let amount = '123456789012345678.90';
  let principal = 'entra-object-1';
  const fetch = vi.fn();
  const client = () =>
    new StreamAssistClient(
      { getAccessToken: async () => 'unused' },
      {
        assistant: { project: 'p', location: 'eu', engine: 'e' },
        identity: 'test',
      },
      fetch as never,
    );
  const snapshot = () =>
    makeCellSnapshot({
      surface: 'excel',
      documentId: 'source-workbook',
      locator: 'Metrics!A1:A2',
      objectId: 'sheet-metrics',
      values: [['Revenue'], [amount]],
      formulas: [[''], ['']],
    });
  const sourceBridge: DocBridge = {
    surface: 'excel',
    documentKey: () => 'source-workbook',
    getCapabilities: () => ({ surface: 'excel', contextKinds: ['range'], actuations: [] }),
    listContext: async () => [],
    resolveContext: async () => [],
    captureCells: snapshot,
    actuate: vi.fn(),
  };
  const sourceSession = new AssistSession(sourceBridge, client(), {
    unit: { connectors: [], surfaceContext: { kind: 'excel' } },
    autoAttach: [],
    compute: async () => {
      throw new Error('Capture needs no compute inference.');
    },
  });
  await collect(
    sourceSession.runAnalysis({ kind: 'capture', range: 'Metrics!A1:A2', headers: true }),
  );
  const artifact = await new ArtifactStore().fromSnapshot(await snapshot());
  const source = createOfficeApplicationPort(
    sourceSession,
    'entra-object-1',
    async () => principal,
  );
  async function destination(surface: 'word' | 'powerpoint' | 'outlook') {
    let text = 'Old generated output';
    let identity = 'entra-object-1';
    let key = `${surface}-document`;
    let uncertain = false;
    const context = async () => ({
      surface,
      documentId: await contentHash({ surface, documentKey: key }),
    });
    const capture = async (): Promise<TextTargetCapture> => ({
      text,
      target: {
        ...(await context()),
        objectId: 'object-1',
        ...(surface === 'powerpoint' ? { parentObjectId: 'slide-1' } : {}),
        locator: `${surface}:target`,
        revision: await contentHash({ text }),
        scope: surface === 'word' ? 'word-content-control' : 'powerpoint-shape-text',
      },
    });
    const kind =
      surface === 'word'
        ? 'fill-content-control'
        : surface === 'powerpoint'
          ? 'set-shape-text'
          : 'create-mail';
    const actuate = vi.fn(async (request: ActuationRequest) => {
      if (uncertain)
        return { ok: false, kind: request.kind, changeId: request.changeId, recoveryPending: true };
      if (surface === 'outlook')
        return {
          ok: true,
          kind: request.kind,
          changeId: request.changeId,
          completion: 'review-opened' as const,
        };
      const current = await capture();
      if (request.params.target?.expectedTextHash !== current.target.revision)
        return {
          ok: false,
          kind: request.kind,
          changeId: request.changeId,
          error: { code: 'target_conflict', message: 'Changed' },
        };
      text = request.params.text!;
      return {
        ok: true,
        kind: request.kind,
        changeId: request.changeId,
        verification: { status: 'verified' as const },
      };
    });
    const bridge: DocBridge = {
      surface,
      documentKey: () => key,
      getCapabilities: () => ({
        surface,
        contextKinds: [],
        actuations: [{ surface, kind, title: kind, reversible: false }],
      }),
      listContext: async () => [],
      resolveContext: async () => [],
      captureTextTarget: capture,
      actuate,
    };
    const session = new AssistSession(bridge, client(), {
      unit: { connectors: [], surfaceContext: { kind: surface } },
      autoAttach: [],
      taskJournal: new TaskJournal(
        new MemoryTaskJournalStore(),
        await contentHash({ surface, key }),
      ),
      durablePrograms: new MemoryDurableProgramStore(),
    });
    return {
      session,
      port: createOfficeApplicationPort(session, 'entra-object-1', async () => identity),
      capture,
      actuate,
      setIdentity: (value: string) => {
        identity = value;
      },
      setDocument: (value: string) => {
        key = value;
      },
      setText: (value: string) => {
        text = value;
      },
      setUncertain: () => {
        uncertain = true;
      },
    };
  }
  const word = await destination('word');
  const powerpoint = await destination('powerpoint');
  const outlook = await destination('outlook');
  const coordinator = new OfficeWorkflowCoordinator();
  coordinator.register('source', source);
  for (const [id, target] of [
    ['word', word],
    ['powerpoint', powerpoint],
    ['outlook', outlook],
  ] as const)
    coordinator.register(id, target.port);
  const metrics = {
    template: 'Revenue: {{revenue}}',
    values: [{ name: 'revenue', row: 0, column: 0 }],
  };
  const plan: OfficeHandoffPlan = {
    version: 1,
    id: 'board-pack',
    sourceApplication: 'source',
    artifactId: artifact.id,
    artifactHash: artifact.hash,
    sourceBinding: await source.binding(),
    steps: [
      {
        id: 'memo',
        application: 'word',
        kind: 'managed-text',
        binding: await word.port.binding(),
        capture: await word.capture(),
        previousGeneratedText: 'Old generated output',
        metrics,
      },
      {
        id: 'deck',
        application: 'powerpoint',
        kind: 'managed-text',
        binding: await powerpoint.port.binding(),
        capture: await powerpoint.capture(),
        previousGeneratedText: 'Old generated output',
        metrics,
      },
      {
        id: 'response',
        application: 'outlook',
        kind: 'outlook-review',
        binding: await outlook.port.binding(),
        subject: 'Month close',
        metrics,
      },
    ],
  };
  return {
    coordinator,
    plan,
    sourceSession,
    source,
    word,
    powerpoint,
    outlook,
    fetch,
    changeSource: () => {
      amount = '999.00';
    },
    changeSourcePrincipal: () => {
      principal = 'entra-object-2';
    },
  };
}

describe('cross-application metric handoffs through real session executors (simulated hosts)', () => {
  it('carries exact source metrics through Word, PowerPoint and an explicitly reviewable Outlook form', async () => {
    const f = await fixture();
    const approve = vi.fn(async () => true);
    const events = await collect(f.coordinator.run(f.plan, { approvePlan: approve }));
    expect(
      events.filter((event) => event.type === 'handoff-step').map((event) => event.status),
    ).toEqual(['verified', 'verified', 'review-required']);
    expect(approve).toHaveBeenCalledTimes(3);
    expect(f.word.actuate.mock.calls[0]![0].params.text).toBe('Revenue: 123456789012345678.90');
    expect(f.outlook.actuate.mock.calls[0]![0].params.mail).toEqual({
      subject: 'Month close',
      body: 'Revenue: 123456789012345678.90',
    });
    expect(f.fetch).not.toHaveBeenCalled();
  });
  it('reopens verified destinations after fresh source/output validation without repeated writes or approvals', async () => {
    const f = await fixture();
    f.plan.steps.pop();
    const approve = vi.fn(async () => true);
    await collect(f.coordinator.run(f.plan, { approvePlan: approve }));
    await collect(f.coordinator.run(f.plan, { approvePlan: approve }));
    expect(f.word.actuate).toHaveBeenCalledOnce();
    expect(f.powerpoint.actuate).toHaveBeenCalledOnce();
    expect(approve).toHaveBeenCalledTimes(2);
    f.word.setText('A human changed the memo');
    await expect(collect(f.coordinator.run(f.plan, { approvePlan: approve }))).rejects.toThrow(
      'fresh source',
    );
  });
  it('blocks a source change during approval at the final pre-dispatch guard', async () => {
    const f = await fixture();
    const events = await collect(
      f.coordinator.run(f.plan, {
        approvePlan: () => {
          f.changeSource();
          return true;
        },
      }),
    );
    expect(events).toContainEqual(
      expect.objectContaining({ type: 'handoff-step', status: 'blocked' }),
    );
    expect(f.word.actuate).not.toHaveBeenCalled();
    expect(f.powerpoint.actuate).not.toHaveBeenCalled();
  });
  it('checks every destination principal before the first effect and again after approval', async () => {
    const f = await fixture();
    f.powerpoint.setIdentity('someone-else');
    await expect(collect(f.coordinator.run(f.plan, { approvePlan: () => true }))).rejects.toThrow(
      'authenticated account changed',
    );
    expect(f.word.actuate).not.toHaveBeenCalled();
    f.powerpoint.setIdentity('entra-object-1');
    await collect(
      f.coordinator.run(f.plan, {
        approvePlan: () => {
          f.word.setIdentity('someone-else');
          return true;
        },
      }),
    );
    expect(f.word.actuate).not.toHaveBeenCalled();
  });
  it('preserves a native target edit and does not continue after its refusal', async () => {
    const f = await fixture();
    await collect(
      f.coordinator.run(f.plan, {
        approvePlan: () => {
          f.word.setText('Manual edit');
          return true;
        },
      }),
    );
    expect(f.word.actuate).toHaveBeenCalledOnce();
    expect((await f.word.capture()).text).toBe('Manual edit');
    expect(f.powerpoint.actuate).not.toHaveBeenCalled();
  });
  it('never replays an uncertain destination even when the caller repeats the same handoff', async () => {
    const f = await fixture();
    f.word.setUncertain();
    await collect(f.coordinator.run(f.plan, { approvePlan: () => true }));
    await expect(collect(f.coordinator.run(f.plan, { approvePlan: () => true }))).rejects.toThrow();
    expect(f.word.actuate).toHaveBeenCalledOnce();
  });
  it('requires approvals and refuses a mismatched artifact fingerprint', async () => {
    const f = await fixture();
    await collect(f.coordinator.run(f.plan));
    expect(f.word.actuate).not.toHaveBeenCalled();
    f.plan.artifactHash = 'sha256:' + 'b'.repeat(64);
    await expect(collect(f.coordinator.run(f.plan, { approvePlan: () => true }))).rejects.toThrow(
      'reviewed content',
    );
  });
  it('refuses missing authenticated identity and switched destination documents', async () => {
    const f = await fixture();
    await expect(
      createOfficeApplicationPort(
        f.sourceSession,
        'entra-object-1',
        async () => undefined,
      ).binding(),
    ).rejects.toThrow('authenticated object ID');
    f.outlook.setDocument('different-item');
    await expect(collect(f.coordinator.run(f.plan, { approvePlan: () => true }))).rejects.toThrow(
      'reviewed destination',
    );
    expect(f.word.actuate).not.toHaveBeenCalled();
  });
  it('pins the composition account even before a port has ever been bound', async () => {
    const f = await fixture();
    const switched = createOfficeApplicationPort(
      f.sourceSession,
      'original-account',
      async () => 'new-account',
    );
    await expect(switched.binding()).rejects.toThrow('authenticated account changed');
    f.changeSourcePrincipal();
    await expect(collect(f.coordinator.run(f.plan, { approvePlan: () => true }))).rejects.toThrow(
      'authenticated account changed',
    );
    expect(f.word.actuate).not.toHaveBeenCalled();
  });
});
