import { describe, expect, it, vi } from 'vitest';
import {
  asChangeId,
  makeCellSnapshot,
  gridForRequest,
  type ActuationRequest,
  type CellValue,
} from '@ge/contracts';
import type { DocBridge } from './bridge.js';
import { RecoveryCoordinator } from './recovery.js';
import { AssistSession } from './assist-session.js';
import { RuntimeHooks } from './hooks.js';
import { completedEffectsExtension, installRuntimeExtensions } from './extensions.js';
import { TriggerRegistry } from '@ge/triggers';
import type { StreamAssistClient } from '@ge/gemini-client';
function fixture() {
  let cells: CellValue[][] = [[1]];
  let persisted: unknown = [];
  let saves = 0;
  let failAt = -1;
  let landed = 0;
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
      makeCellSnapshot({
        surface: 'excel',
        documentId: 'doc',
        objectId: 'sheet-1',
        locator: locator.includes('!') ? locator : `Sheet1!${locator}`,
        values: cells,
      }),
    recoveryStorage: {
      load: async () => structuredClone(persisted),
      save: async (_owner, value) => {
        if (++saves === failAt) throw new Error('checkpoint failed');
        persisted = structuredClone(value);
      },
    },
    actuate: async (req) => {
      const before = await bridge.captureCells!(req.params.target!.range!);
      if (req.preconditions?.some((p) => p.hash !== before.hash))
        return {
          ok: false,
          kind: req.kind,
          changeId: req.changeId,
          error: { code: 'stale_target', message: 'Cells changed' },
        };
      landed++;
      cells = structuredClone(gridForRequest(req));
      const after = await bridge.captureCells!(req.params.target!.range!);
      return {
        ok: true,
        kind: req.kind,
        changeId: req.changeId,
        verification: { status: 'verified', beforeHash: before.hash, afterHash: after.hash },
      };
    },
  };
  const request = (value = 2): ActuationRequest => ({
    changeId: asChangeId(crypto.randomUUID()),
    surface: 'excel',
    kind: 'write-cells',
    params: { target: { range: 'A1' }, cellValues: [[value]] },
  });
  return {
    bridge,
    request,
    failSave(n: number) {
      failAt = n;
    },
    edit(value: number) {
      cells = [[value]];
    },
    cells: () => cells,
    landed: () => landed,
    saved: () => structuredClone(persisted),
  };
}
describe('durable recovery and verified writes', () => {
  it('pins the destination before approval and blocks a stale destination', async () => {
    const f = fixture();
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    const request = await recovery.prepare(f.request());
    expect(request.params.target?.range).toBe('Sheet1!A1');
    f.edit(99);
    const result = await recovery.execute(request, () => f.bridge.actuate(request));
    expect(result.ok).toBe(false);
    expect(f.landed()).toBe(0);
    expect(f.cells()).toEqual([[99]]);
  });
  it('fails closed when the intent cannot be checkpointed', async () => {
    const f = fixture();
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    const request = await recovery.prepare(f.request());
    f.failSave(1);
    await expect(recovery.execute(request, () => f.bridge.actuate(request))).rejects.toThrow(
      'checkpoint',
    );
    expect(f.landed()).toBe(0);
  });
  it('preserves a landed write after a receipt save failure and recovers it after reload', async () => {
    const f = fixture();
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    const request = await recovery.prepare(f.request());
    f.failSave(2);
    const result = await recovery.execute(request, () => f.bridge.actuate(request));
    expect(result).toMatchObject({ ok: true, recoveryPending: true });
    const restarted = new RecoveryCoordinator(f.bridge, 'owner');
    expect(await restarted.inspect()).toMatchObject([{ state: 'applied', canUndo: true }]);
    expect(f.landed()).toBe(1);
  });
  it('reconciles an interrupted intent, requires a new request and blocks undo after editing', async () => {
    const f = fixture();
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    const request = await recovery.prepare(f.request());
    await recovery.execute(request, async () => {
      throw new Error('connection interrupted');
    });
    const resumed = new RecoveryCoordinator(f.bridge, 'owner');
    expect(await resumed.inspect()).toMatchObject([{ state: 'not-applied', canResume: true }]);
    const retry = await resumed.prepare(await resumed.request(request.changeId, false));
    expect(retry.changeId).not.toBe(request.changeId);
    expect(f.landed()).toBe(0);
    await resumed.execute(retry, () => f.bridge.actuate(retry));
    f.edit(500);
    await expect(resumed.request(retry.changeId, true)).rejects.toThrow('cannot be replayed');
    expect(f.cells()).toEqual([[500]]);
  });
  it('executes undo through a fresh plan approval and the same effect hooks', async () => {
    const f = fixture();
    const hooks = new RuntimeHooks();
    const before = vi.fn();
    hooks.register({ id: 'test-before', on: 'effect:before', mode: 'guard', handle: before });
    const session = new AssistSession(f.bridge, {} as StreamAssistClient, {
      unit: { connectors: [], surfaceContext: { kind: 'excel' } },
      hooks,
      recoveryOwner: 'owner',
    });
    const request = f.request();
    await session.apply(request.kind, request.params, request.changeId);
    const approvals = vi.fn(() => false);
    for await (const _ of session.runAnalysis(
      { kind: 'undo', id: request.changeId },
      { approvePlan: approvals },
    )) {
      /* consume */
    }
    expect(approvals).toHaveBeenCalledOnce();
    expect(f.cells()).toEqual([[2]]);
    expect(f.landed()).toBe(1);
    for await (const _ of session.runAnalysis(
      { kind: 'undo', id: request.changeId },
      { approvePlan: () => true },
    )) {
      /* consume */
    }
    expect(f.cells()).toEqual([[1]]);
    expect(before).toHaveBeenCalledTimes(2);
  });
  it('does not mark a readback mismatch as complete or enable automatic undo', async () => {
    const f = fixture();
    const original = f.bridge.actuate;
    f.bridge.actuate = async (r) => ({
      ...(await original(r)),
      verification: { status: 'mismatch', afterHash: 'untrusted' },
    });
    const hooks = new RuntimeHooks();
    installRuntimeExtensions([completedEffectsExtension], {
      hooks,
      triggers: new TriggerRegistry(),
    });
    const session = new AssistSession(f.bridge, {} as StreamAssistClient, {
      unit: { connectors: [], surfaceContext: { kind: 'excel' } },
      hooks,
    });
    const req = f.request();
    await expect(session.apply(req.kind, req.params, req.changeId)).rejects.toThrow(
      'did not complete',
    );
    expect(f.landed()).toBe(1);
    expect(session.recovery.list()[0]).toMatchObject({ state: 'uncertain', canUndo: false });
  });
  it('rejects foreign or malformed persisted receipts', async () => {
    const f = fixture();
    f.bridge.recoveryStorage!.load = async () => [{ version: 999 }];
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    await expect(recovery.inspect()).rejects.toThrow();
    expect(f.landed()).toBe(0);
  });
});
