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

/** Minimal host range simulator for footprint checks; no model, Office network or formula eval. */
function spatialFixture() {
  const f = fixture();
  let documentId = 'doc';
  let value = 1;
  f.bridge.captureCells = async (locator) => {
    const bang = locator.lastIndexOf('!');
    const sheet = bang < 0 ? 'Main' : locator.slice(0, bang).replace(/^'(.*)'$/, '$1');
    const address = locator
      .slice(bang + 1)
      .replace(/\$/g, '')
      .toUpperCase();
    const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(address)!;
    const column = (s: string) => [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
    const columns = column(match[3] ?? match[1]!) - column(match[1]!) + 1;
    const rows = Number(match[4] ?? match[2]) - Number(match[2]) + 1;
    return makeCellSnapshot({
      surface: 'excel',
      documentId,
      objectId: `sheet:${sheet === 'Renamed Main' ? 'main' : sheet.toLowerCase()}`,
      locator: `${sheet}!${address}`,
      values: Array.from({ length: rows }, () => Array.from({ length: columns }, () => value)),
    });
  };
  return {
    ...f,
    request(
      range = 'Main!B2',
      values: CellValue[][] = [
        [2, 3],
        [4, 5],
      ],
    ): ActuationRequest {
      return { ...f.request(), params: { target: { range }, cellValues: values } };
    },
    switchDocument(id: string) {
      documentId = id;
    },
    edit(valueToWrite: number) {
      value = valueToWrite;
    },
  };
}

describe('unresolved write footprints', () => {
  it.each(['Main!$B$2', "'Main'!C3", 'Main!A1:B2', 'Main!C3:D4', "'Renamed Main'!B2"])(
    'blocks a fresh receipt for overlapping destination %s after reload',
    async (range) => {
      const f = spatialFixture();
      const first = new RecoveryCoordinator(f.bridge, 'owner');
      const request = await first.prepare(f.request());
      await first.execute(request, async () => {
        throw new Error('host disconnected');
      });
      const restarted = new RecoveryCoordinator(f.bridge, 'owner');
      const values = range.includes(':')
        ? [
            [9, 9],
            [9, 9],
          ]
        : [[9]];
      await expect(restarted.prepare(f.request(range, values))).rejects.toThrow(
        'unresolved write overlaps',
      );
      expect(f.landed()).toBe(0);
    },
  );

  it.each(['Main!A1', 'Main!D4', 'Other!B2'])(
    'allows an unrelated destination %s while another receipt is unresolved',
    async (range) => {
      const f = spatialFixture();
      const recovery = new RecoveryCoordinator(f.bridge, 'owner');
      const request = await recovery.prepare(f.request());
      await recovery.execute(request, async () => {
        throw new Error('host disconnected');
      });
      const next = await recovery.prepare(f.request(range, [[9]]));
      const actuate = vi.fn(async () => ({ ok: false, kind: next.kind, changeId: next.changeId }));
      await recovery.execute(next, actuate);
      expect(actuate).toHaveBeenCalledOnce();
    },
  );

  it('does not block the same address in a different document', async () => {
    const f = spatialFixture();
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    const request = await recovery.prepare(f.request());
    await recovery.execute(request, async () => {
      throw new Error('host disconnected');
    });
    f.switchDocument('other-doc');
    await expect(recovery.prepare(f.request())).resolves.toMatchObject({ kind: 'write-cells' });
  });

  it('blocks a durable applying intent left by a failed receipt checkpoint', async () => {
    const f = fixture();
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    const request = await recovery.prepare(f.request());
    f.failSave(2);
    await recovery.execute(request, () => f.bridge.actuate(request));
    const restarted = new RecoveryCoordinator(f.bridge, 'owner');
    await expect(restarted.prepare(f.request(3))).rejects.toThrow('unresolved write overlaps');
    expect(f.landed()).toBe(1);
    // Inspection reconciles the intent; a verified result can then use the explicit undo route.
    const undo = await restarted.prepare(await restarted.request(request.changeId, true));
    await restarted.execute(undo, () => f.bridge.actuate(undo));
    expect(f.cells()).toEqual([[1]]);
  });

  it('keeps a conflicted receipt blocking overlapping writes after inspection', async () => {
    const f = spatialFixture();
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    const request = await recovery.prepare(f.request());
    await recovery.execute(request, async () => {
      throw new Error('host disconnected');
    });
    f.edit(500);
    expect(await recovery.inspect()).toMatchObject([{ state: 'conflict', canForget: false }]);
    await expect(recovery.prepare(f.request())).rejects.toThrow('unresolved write overlaps');
    await expect(recovery.request(request.changeId, false)).rejects.toThrow('cannot be replayed');
    await expect(recovery.forget(request.changeId)).rejects.toThrow('Unresolved receipts');
  });

  it('allows a new reviewed write after a verified historical write was manually changed', async () => {
    const f = fixture();
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    const original = await recovery.prepare(f.request());
    await recovery.execute(original, () => f.bridge.actuate(original));
    f.edit(99);
    expect(await recovery.inspect()).toMatchObject([
      { state: 'conflict', canUndo: false, canResume: false, canForget: true },
    ]);
    await expect(recovery.request(original.changeId, true)).rejects.toThrow('cannot be replayed');
    // A fresh coordinator reloads the historical conflict, then permits independently reviewed work.
    const next = new RecoveryCoordinator(f.bridge, 'owner');
    const request = await next.prepare(f.request(3));
    expect(await next.execute(request, () => f.bridge.actuate(request))).toMatchObject({
      ok: true,
      verification: { status: 'verified' },
    });
    expect(f.cells()).toEqual([[3]]);
    await next.forget(original.changeId);
    expect(next.list()).toHaveLength(1);
    expect(next.list()[0]?.id).toBe(request.changeId);
    expect(f.landed()).toBe(2);
    expect(f.cells()).toEqual([[3]]);
  });

  it.each([
    { recoveryPending: true },
    { error: { code: 'outcome_unknown', message: 'The host outcome is unknown.' } },
    { verification: { status: 'unknown' as const } },
    { verification: { status: 'mismatch' as const } },
  ])('persists a failed result with uncertainty as unresolved: %j', async (outcome) => {
    const f = fixture();
    const recovery = new RecoveryCoordinator(f.bridge, 'owner');
    const request = await recovery.prepare(f.request());
    await recovery.execute(request, async () => ({
      ok: false,
      kind: request.kind,
      changeId: request.changeId,
      ...outcome,
    }));
    expect(recovery.list()).toMatchObject([{ state: 'uncertain', canResume: false }]);
    await expect(recovery.prepare(f.request(3))).rejects.toThrow('unresolved write overlaps');
  });

  it('rechecks storage under the execution lock when two panes prepared before either wrote', async () => {
    const f = spatialFixture();
    const first = new RecoveryCoordinator(f.bridge, 'owner');
    const second = new RecoveryCoordinator(f.bridge, 'owner');
    const a = await first.prepare(f.request());
    const b = await second.prepare(f.request('Main!C3', [[9]]));
    let started!: () => void;
    const entered = new Promise<void>((resolve) => {
      started = resolve;
    });
    let release!: () => void;
    const interrupted = new Promise<void>((resolve) => {
      release = resolve;
    });
    const firstRun = first.execute(a, async () => {
      started();
      await interrupted;
      throw new Error('Office disconnected after dispatch');
    });
    await entered;
    const actuate = vi.fn(async () => ({ ok: true, kind: b.kind, changeId: b.changeId }));
    const secondRun = second.execute(b, actuate).catch((error: unknown) => error);
    release();
    expect(await firstRun).toMatchObject({ recoveryPending: true });
    expect(await secondRun).toEqual(
      expect.objectContaining({ message: expect.stringContaining('unresolved write overlaps') }),
    );
    expect(actuate).not.toHaveBeenCalled();
    await expect(second.prepare(b)).rejects.toThrow('unresolved write overlaps');
  });

  it('revalidates a recovery parent so two prepared resumes cannot both execute', async () => {
    const f = fixture();
    const first = new RecoveryCoordinator(f.bridge, 'owner');
    const interrupted = await first.prepare(f.request());
    await first.execute(interrupted, async () => {
      throw new Error('host disconnected');
    });
    const second = new RecoveryCoordinator(f.bridge, 'owner');
    const a = await first.prepare(await first.request(interrupted.changeId, false));
    const b = await second.prepare(await second.request(interrupted.changeId, false));
    await first.execute(a, () => f.bridge.actuate(a));
    const actuate = vi.fn(() => f.bridge.actuate(b));
    await expect(second.execute(b, actuate)).rejects.toThrow('recovery receipt changed');
    expect(actuate).not.toHaveBeenCalled();
    expect(f.landed()).toBe(1);
  });
});

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
