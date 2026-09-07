import { describe, expect, it, vi } from 'vitest';
import {
  asChangeId,
  makeCellSnapshot,
  type ActuationRequest,
  type ActuationResult,
  type CapabilityManifest,
} from '@ge/contracts';
import type { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession } from './assist-session.js';
import type { DocBridge } from './bridge.js';
import { admitActuationRequest } from './actuation-admission.js';
import { RecoveryCoordinator } from './recovery.js';

const manifest: CapabilityManifest = {
  surface: 'word',
  contextKinds: [],
  actuations: [{ kind: 'tracked-change', surface: 'word', title: 'Edit', reversible: true }],
};
const request: ActuationRequest = {
  surface: 'word',
  kind: 'tracked-change',
  changeId: asChangeId('expected'),
  params: { text: 'Changed' },
};
function fixture(
  options: { disabled?: boolean; raw?: unknown; throws?: boolean; revoke?: boolean } = {},
) {
  let reads = 0;
  const actuate = vi.fn(async (effect: ActuationRequest): Promise<ActuationResult> => {
    if (options.throws) throw new Error('Host disconnected after dispatch');
    return (
      Object.hasOwn(options, 'raw')
        ? options.raw
        : { ok: true, kind: effect.kind, changeId: effect.changeId }
    ) as ActuationResult;
  });
  const bridge: DocBridge = {
    surface: 'word',
    getCapabilities: () => {
      reads++;
      return options.revoke && reads > 1 ? { ...manifest, actuations: [] } : manifest;
    },
    listContext: async () => [],
    resolveContext: async () => [],
    actuate,
  };
  const session = new AssistSession(bridge, {} as StreamAssistClient, {
    unit: { connectors: [], surfaceContext: { kind: 'word' } },
    ...(options.disabled
      ? { capabilityFilter: (value: CapabilityManifest) => ({ ...value, actuations: [] }) }
      : {}),
  });
  return { session, actuate };
}

describe('one actuation boundary for proposals, commands and recovery', () => {
  it.each([{ disabled: true }, { revoke: true }])(
    'honors effective capability admission and re-admission',
    async (options) => {
      const { session, actuate } = fixture(options);
      const result = await session.apply(request.kind, request.params, request.changeId);
      expect(result).toMatchObject({ ok: false, error: { code: 'capability_unavailable' } });
      expect(actuate).not.toHaveBeenCalled();
      session.dispose();
    },
  );

  it('validates requests and surface identity before execution', () => {
    expect(
      admitActuationRequest({ ...request, surface: 'excel' }, 'word', manifest).rejection,
    ).toMatchObject({ ok: false, error: { code: 'surface_mismatch' } });
    expect(() =>
      admitActuationRequest({ ...request, params: { text: 12 } }, 'word', manifest),
    ).toThrow();
  });

  it.each([
    { raw: undefined },
    {
      raw: {
        ok: true,
        changeId: 'wrong',
        kind: 'write-cells',
        verification: { status: 'verified' },
      },
    },
    { throws: true },
  ])('records uncertain host outcomes against the actual requested effect', async (options) => {
    const { session, actuate } = fixture(options);
    const result = await session.apply(request.kind, request.params, request.changeId);
    expect(result).toMatchObject({
      ok: false,
      changeId: request.changeId,
      kind: request.kind,
      recoveryPending: true,
      error: { code: 'outcome_unknown' },
    });
    expect(actuate).toHaveBeenCalledTimes(1);
    expect(session.executions.list().at(-1)).toMatchObject({
      status: 'incomplete',
      effects: [
        {
          changeId: request.changeId,
          kind: request.kind,
          ok: false,
          errorCode: 'outcome_unknown',
        },
      ],
    });
    session.dispose();
  });

  it('validates receipts before the durable recovery state can classify an effect', async () => {
    const bridge: DocBridge = {
      surface: 'excel',
      getCapabilities: () => ({ surface: 'excel', contextKinds: [], actuations: [] }),
      listContext: async () => [],
      resolveContext: async () => [],
      actuate: async () => {
        throw new Error('unused');
      },
      captureCells: async (locator) =>
        makeCellSnapshot({ surface: 'excel', documentId: 'doc', locator, values: [['before']] }),
    };
    const recovery = new RecoveryCoordinator(bridge, 'owner');
    const prepared = await recovery.prepare({
      kind: 'write-cells',
      surface: 'excel',
      changeId: asChangeId('actual'),
      params: { target: { range: 'Sheet1!A1' }, cellValues: [['after']] },
    });
    const receipt = await recovery.execute(prepared, async () => ({
      ok: true,
      kind: 'write-cells',
      changeId: asChangeId('foreign'),
      verification: { status: 'verified', afterHash: 'unrelated' },
    }));
    expect(receipt.error?.code).toBe('outcome_unknown');
    expect(recovery.list()).toEqual([
      expect.objectContaining({
        id: 'actual',
        state: 'uncertain',
        canResume: false,
        canUndo: false,
      }),
    ]);
  });
});
