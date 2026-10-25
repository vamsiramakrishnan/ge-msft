import { describe, expect, it, vi } from 'vitest';
import { contentHash, type ActuationRequest, type TextTargetCapture } from '@ge/contracts';
import type { StreamAssistClient } from '@ge/gemini-client';
import {
  AssistSession,
  MemoryTaskJournalStore,
  TaskJournal,
  MemoryDurableProgramStore,
  type DocBridge,
} from '@ge/runtime';
import { PanelController } from './controller.js';

function held() {
  let release!: () => void;
  const waiting = new Promise<void>((resolve) => {
    release = resolve;
  });
  return { waiting, release };
}
async function fixture() {
  let key = 'document-1';
  let account = 'account-1';
  let text = 'Original text';
  let pause: ReturnType<typeof held> | undefined;
  const capture = async (): Promise<TextTargetCapture> => ({
    text,
    target: {
      surface: 'word',
      documentId: await contentHash({ surface: 'word', documentKey: key }),
      objectId: '7',
      scope: 'word-content-control',
      locator: 'word:cc:7',
      revision: await contentHash({ text }),
    },
  });
  const actuate = vi.fn(async (request: ActuationRequest) => {
    if (pause) await pause.waiting;
    text = request.params.text!;
    return {
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      verification: { status: 'verified' as const, afterHash: await contentHash({ text }) },
    };
  });
  const bridge: DocBridge = {
    surface: 'word',
    documentKey: () => key,
    getCapabilities: () => ({
      surface: 'word',
      contextKinds: [],
      actuations: [
        {
          kind: 'fill-content-control',
          surface: 'word',
          title: 'Fill managed text',
          reversible: true,
        },
      ],
    }),
    listContext: async () => [],
    resolveContext: async () => [],
    captureTextTarget: capture,
    actuate,
  };
  const programs = new MemoryDurableProgramStore();
  const journalStore = new MemoryTaskJournalStore();
  const session = new AssistSession(bridge, {} as StreamAssistClient, {
    unit: { connectors: [], surfaceContext: { kind: 'word' } },
    autoAttach: [],
    durablePrograms: programs,
    validatePrincipal: async () => {
      if (account !== 'account-1') throw new Error('Authenticated account changed.');
    },
    taskJournalFactory: async (doc) =>
      new TaskJournal(journalStore, await contentHash({ account: 'account-1', doc })),
  });
  const controller = new PanelController(session, bridge);
  const source =
    '/fill-content-control contentControlId="7" text="Reviewed close"\nfinish when=verified';
  const save = () => controller.saveLiteralOfficeTask('close-memo', source, true);
  return {
    controller,
    session,
    bridge,
    programs,
    actuate,
    source,
    save,
    switchDocument: () => {
      key = 'document-2';
    },
    switchAccount: () => {
      account = 'account-2';
    },
    pause: () => {
      pause = held();
      return pause;
    },
  };
}
async function waitFor(test: () => boolean): Promise<void> {
  for (let attempt = 0; attempt < 100; attempt++) {
    if (test()) return;
    await new Promise((resolve) => setTimeout(resolve, 1));
  }
  throw new Error('Expected controller state was not reached.');
}

describe('saved-task pane through the real session executor (simulated native host)', () => {
  it('saves and inspects literal commands without execution, then requires a fresh ordinary plan approval', async () => {
    const f = await fixture();
    expect(f.controller.getState().savedOfficeTasks?.loaded).toBe(false);
    await f.save();
    expect(f.actuate).not.toHaveBeenCalled();
    expect(f.controller.getState().savedOfficeTasks?.error).toBeUndefined();
    expect(f.controller.getState().savedOfficeTasks?.selected?.canResume).toBe(true);
    const run = f.controller.resumeSavedOfficeTask('close-memo');
    await waitFor(() => Boolean(f.controller.getState().pendingPlan));
    expect(f.actuate).not.toHaveBeenCalled();
    expect(f.controller.getState().pendingPlan?.effects[0]?.request.kind).toBe(
      'fill-content-control',
    );
    f.controller.approvePlan();
    await run;
    expect(f.actuate).toHaveBeenCalledOnce();
    expect(f.controller.getState().savedOfficeTasks?.items[0]).toMatchObject({
      state: 'verified',
      verifiedStages: 1,
    });
    expect(f.controller.getState().messages.at(-1)?.changes).toHaveLength(1);
    f.session.dispose();
  });
  it('does not save without literal attestation and never substitutes the model command route', async () => {
    const f = await fixture();
    const model = vi.spyOn(f.session, 'runCommands');
    await f.controller.saveLiteralOfficeTask('close-memo', f.source, false);
    expect(f.controller.getState().savedOfficeTasks?.error).toContain('Confirm');
    expect(await f.session.listDurablePrograms()).toEqual([]);
    expect(f.actuate).not.toHaveBeenCalled();
    expect(model).not.toHaveBeenCalled();
    f.session.dispose();
  });
  it('blocks concurrent pane actions while inspection owns the controller and drops stale document results', async () => {
    const f = await fixture();
    await f.save();
    const gate = held();
    const assess = f.session.assessDurableProgram.bind(f.session);
    const inspect = vi.spyOn(f.session, 'assessDurableProgram').mockImplementation(async (id) => {
      const result = await assess(id);
      await gate.waiting;
      return result;
    });
    const pending = f.controller.inspectSavedOfficeTask('close-memo');
    await waitFor(() => f.controller.getState().busy);
    await f.controller.inspectSavedOfficeTask('close-memo');
    expect(inspect).toHaveBeenCalledOnce();
    f.switchDocument();
    gate.release();
    await pending;
    expect(f.controller.getState().savedOfficeTasks?.selected).toBeUndefined();
    expect(f.controller.getState().savedOfficeTasks?.items).toEqual([]);
    expect(f.controller.getState().savedOfficeTasks?.error).toContain('document changed');
    expect(f.controller.getState().busy).toBe(false);
    f.session.dispose();
  });
  it('refuses an account switch and clears previously inspected task bodies', async () => {
    const f = await fixture();
    await f.save();
    expect(f.controller.getState().savedOfficeTasks?.selected).toBeDefined();
    f.switchAccount();
    await f.controller.refreshSavedOfficeTasks();
    expect(f.controller.getState().savedOfficeTasks?.selected).toBeUndefined();
    expect(f.controller.getState().savedOfficeTasks?.error).toContain('account changed');
    expect(f.actuate).not.toHaveBeenCalled();
    f.session.dispose();
  });
  it('preserves a landed native receipt when cancellation happens after dispatch', async () => {
    const f = await fixture();
    await f.save();
    const gate = f.pause();
    const run = f.controller.resumeSavedOfficeTask('close-memo');
    await waitFor(() => Boolean(f.controller.getState().pendingPlan));
    f.controller.approvePlan();
    await waitFor(() => f.actuate.mock.calls.length === 1);
    f.controller.cancel();
    gate.release();
    await run;
    const reply = f.controller.getState().messages.at(-1)!;
    expect(reply.cancelled).toBe(true);
    expect(reply.changes).toHaveLength(1);
    expect(f.actuate).toHaveBeenCalledOnce();
    expect(f.controller.getState().busy).toBe(false);
    f.session.dispose();
  });
  it('hides completed tasks without allowing their IDs or effects to be replayed', async () => {
    const f = await fixture();
    await f.save();
    const run = f.controller.resumeSavedOfficeTask('close-memo');
    await waitFor(() => Boolean(f.controller.getState().pendingPlan));
    f.controller.approvePlan();
    await run;
    await f.controller.archiveSavedOfficeTask('close-memo');
    expect(f.controller.getState().savedOfficeTasks?.items).toEqual([]);
    expect((await f.session.inspectDurableProgram('close-memo'))?.archivedAt).toBeDefined();
    await f.controller.saveLiteralOfficeTask('close-memo', f.source, true);
    expect(f.controller.getState().savedOfficeTasks?.error).toBeDefined();
    expect(f.actuate).toHaveBeenCalledOnce();
    f.session.dispose();
  });
});
