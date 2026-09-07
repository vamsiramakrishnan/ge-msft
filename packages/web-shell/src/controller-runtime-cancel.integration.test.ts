import { describe, expect, it, vi } from 'vitest';
import {
  asChangeId,
  assessActuationResult,
  makeCellSnapshot,
  type ActuationRequest,
} from '@ge/contracts';
import type { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession, type DocBridge } from '@ge/runtime';
import { PanelController } from './controller.js';

function gate() {
  let enter: () => void = () => {};
  let release: () => void = () => {};
  return {
    entered: new Promise<void>((resolve) => {
      enter = resolve;
    }),
    waiting: new Promise<void>((resolve) => {
      release = resolve;
    }),
    enter: () => enter(),
    release: () => release(),
  };
}

/** Real controller/session/hooks/recovery pipeline. Only host operations are deterministic adapters. */
function fixture(cells = false) {
  const surface = cells ? 'excel' : 'word';
  const kind = cells ? 'write-cells' : 'tracked-change';
  const actuate = vi.fn(async (request: ActuationRequest) => ({
    kind: request.kind,
    changeId: request.changeId,
    ok: true,
    verification: { status: 'verified' as const },
  }));
  const bridge: DocBridge = {
    surface,
    getCapabilities: () => ({
      surface,
      contextKinds: [cells ? 'range' : 'selection'],
      actuations: [{ kind, surface, title: 'Write', reversible: true }],
    }),
    listContext: async () => [],
    resolveContext: async () => [],
    ...(cells
      ? {
          captureCells: async (locator: string) =>
            makeCellSnapshot({ surface: 'excel', documentId: 'doc', locator, values: [[0]] }),
        }
      : {}),
    actuate,
  };
  const session = new AssistSession(bridge, {} as StreamAssistClient, {
    unit: { connectors: [], surfaceContext: { kind: surface } },
  });
  const controller = new PanelController(session, { listContext: async () => [] });
  const proposal = controller.propose(
    kind,
    cells
      ? { target: { range: 'Output!A1' }, cellValues: [['42']] }
      : { text: 'Reviewed replacement' },
    'Apply reviewed change',
  );
  return { bridge, session, controller, proposal, actuate };
}

describe('proposal cancellation through the real runtime', () => {
  it.each(['message:received', 'effect:before'] as const)(
    'cancels while %s is waiting and never dispatches the host write',
    async (phase) => {
      const f = fixture();
      const held = gate();
      f.session.hooks.register({
        id: `pause-${phase}`,
        on: phase,
        mode: 'guard',
        handle: async () => {
          held.enter();
          await held.waiting;
        },
      });
      const pending = f.controller.applyProposal(f.proposal.changeId);
      await held.entered;
      f.controller.cancel();
      held.release();
      await pending;
      expect(f.actuate).not.toHaveBeenCalled();
      expect(f.controller.getState().busy).toBe(false);
      expect(f.controller.getState().proposals[0]?.status).not.toBe('applying');
      expect(f.session.executions.list().at(-1)?.status).toBe('cancelled');
      f.session.dispose();
    },
  );

  it('cancels while recovery captures the destination before dispatch', async () => {
    const f = fixture(true);
    const held = gate();
    const capture = f.bridge.captureCells!;
    f.bridge.captureCells = async (locator) => {
      held.enter();
      await held.waiting;
      return capture(locator);
    };
    const pending = f.controller.applyProposal(f.proposal.changeId);
    await held.entered;
    f.controller.cancel();
    held.release();
    await pending;
    expect(f.actuate).not.toHaveBeenCalled();
    expect(f.controller.getState().busy).toBe(false);
    expect(f.session.executions.list().at(-1)?.status).toBe('cancelled');
    f.session.dispose();
  });

  it('retains the actual verified receipt when cancellation happens after host dispatch', async () => {
    const f = fixture();
    const held = gate();
    f.actuate.mockImplementationOnce(async (request) => {
      held.enter();
      await held.waiting;
      return {
        kind: request.kind,
        changeId: request.changeId,
        ok: true,
        verification: { status: 'verified' },
      };
    });
    const pending = f.controller.applyProposal(f.proposal.changeId);
    await held.entered;
    f.controller.cancel();
    held.release();
    await pending;
    expect(f.actuate).toHaveBeenCalledTimes(1);
    expect(f.controller.getState().proposals[0]?.status).toBe('applied');
    expect(assessActuationResult(f.controller.getState().changes[0]!)).toBe('verified');
    expect(f.session.executions.list().at(-1)).toMatchObject({
      status: 'cancelled',
      effects: [{ changeId: f.proposal.changeId, ok: true }],
    });
    expect(f.controller.getState().busy).toBe(false);
    f.session.dispose();
  });

  it('records a known non-write when cancelled while the applying intent is being saved', async () => {
    const f = fixture(true);
    const held = gate();
    let persisted: unknown = [];
    let firstSave = true;
    f.bridge.recoveryStorage = {
      load: async () => structuredClone(persisted),
      save: async (_owner, value) => {
        persisted = structuredClone(value);
        if (firstSave) {
          firstSave = false;
          held.enter();
          await held.waiting;
        }
      },
    };
    const pending = f.controller.applyProposal(f.proposal.changeId);
    await held.entered;
    expect(f.session.recovery.list()[0]?.state).toBe('applying');
    f.controller.cancel();
    held.release();
    await pending;
    expect(f.actuate).not.toHaveBeenCalled();
    expect(f.controller.getState().busy).toBe(false);
    expect(f.controller.getState().changes[0]).toMatchObject({
      ok: false,
      error: { code: 'cancelled' },
    });
    expect(assessActuationResult(f.controller.getState().changes[0]!)).toBe('failed');
    expect(f.session.recovery.list()[0]).toMatchObject({
      id: f.proposal.changeId,
      state: 'not-applied',
    });
    expect(persisted).toMatchObject([{ id: f.proposal.changeId, state: 'not-applied' }]);
    // A known non-write must not create an uncertain-overlap block on a freshly reviewed effect.
    const nextId = asChangeId('reviewed-after-cancellation');
    await expect(
      f.session.recovery.prepare({
        surface: 'excel',
        kind: 'write-cells',
        changeId: nextId,
        params: { target: { range: 'Output!A1' }, cellValues: [['42']] },
      }),
    ).resolves.toMatchObject({ changeId: nextId });
    expect(f.actuate).not.toHaveBeenCalled();
    f.session.dispose();
  });
});
