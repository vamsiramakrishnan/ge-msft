import { describe, expect, it, vi } from 'vitest';
import { asChangeId, type ActuationRequest, type ActuationResult } from '@ge/contracts';
import {
  createBridgeDispatch as createDispatch,
  type BridgeActuationHandlers,
} from './bridge-dispatch.js';

function createBridgeDispatch<Host>(surface: 'excel', handlers: BridgeActuationHandlers<Host>) {
  return createDispatch(surface, handlers, { provenance: 'reported' });
}

function request(params: ActuationRequest['params'] = {}): ActuationRequest {
  return { changeId: asChangeId('dispatch-1'), kind: 'write-cells', surface: 'excel', params };
}

describe('bridge dispatch boundary', () => {
  it('derives immutable discovery from actual own handlers and snapshots the table', async () => {
    const handler = vi.fn(async (_host: object, req: ActuationRequest) => ({
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
    }));
    const handlers = { 'write-cells': handler };
    const dispatcher = createBridgeDispatch('excel', handlers);
    handlers['write-cells'] = vi.fn();
    expect(dispatcher.handledActuations).toEqual(['write-cells']);
    expect(Object.isFrozen(dispatcher.handledActuations)).toBe(true);
    expect(Object.isFrozen(dispatcher)).toBe(true);
    await dispatcher.dispatch({}, request());
    expect(handler).toHaveBeenCalledOnce();
    expect(handlers['write-cells']).not.toHaveBeenCalled();
  });

  it('does not promote inherited handlers into execution authority', async () => {
    const inherited = vi.fn();
    const handlers = Object.create({ 'write-cells': inherited }) as BridgeActuationHandlers<object>;
    const dispatcher = createBridgeDispatch('excel', handlers);
    expect(dispatcher.handledActuations).toEqual([]);
    expect(await dispatcher.dispatch({}, request())).toMatchObject({
      ok: false,
      error: { code: 'unsupported' },
    });
    expect(inherited).not.toHaveBeenCalled();
  });

  it('rejects malformed handler registrations at construction', () => {
    expect(() => createBridgeDispatch('excel', { 'write-cells': undefined })).toThrow(
      'must be a function',
    );
    expect(() => createBridgeDispatch('excel', { imaginary: vi.fn() } as never)).toThrow();
  });

  it('rejects a valid request for another surface before host access', async () => {
    const host = {};
    const handler = vi.fn();
    const dispatcher = createBridgeDispatch('excel', { 'write-cells': handler });
    const result = await dispatcher.dispatch(host, { ...request(), surface: 'word' });
    expect(result).toMatchObject({ ok: false, error: { code: 'surface_mismatch' } });
    expect(handler).not.toHaveBeenCalled();
  });

  it.each([
    { cellValues: [] },
    { cellValues: [[1]], cellFormulas: [['=SUM(A1:A2)', '']] },
    { text: 42 },
  ])('rejects schema-invalid parameters before host access: %j', async (params) => {
    const handler = vi.fn();
    const dispatcher = createBridgeDispatch('excel', { 'write-cells': handler });
    const result = await dispatcher.dispatch({}, request(params as ActuationRequest['params']));
    expect(result).toMatchObject({
      ok: false,
      changeId: 'dispatch-1',
      kind: 'write-cells',
      error: { code: 'invalid_request' },
    });
    expect(handler).not.toHaveBeenCalled();
  });

  it.each([undefined, null, {}, { kind: 'invented', changeId: 'valid-id' }])(
    'throws on an invalid envelope instead of fabricating an outcome: %j',
    async (input) => {
      const handler = vi.fn();
      const dispatcher = createBridgeDispatch('excel', { 'write-cells': handler });
      await expect(dispatcher.dispatch({}, input as ActuationRequest)).rejects.toThrow();
      expect(handler).not.toHaveBeenCalled();
    },
  );

  it('passes validated request evidence to the host and preserves its entire outcome', async () => {
    const input = request({ target: { range: 'Sheet!A1' }, cellValues: [[42]] });
    input.preconditions = [
      {
        surface: 'excel',
        documentId: 'doc',
        locator: 'Sheet!A1',
        hash: `sha256:${'1'.repeat(64)}`,
      },
    ];
    input.provenance = {
      agentId: 'test',
      identity: 'user@example.com',
      timestamp: '2026-09-07T00:00:00Z',
      contentHash: 'hash',
      sources: [],
    };
    const result: ActuationResult = {
      ok: true,
      kind: input.kind,
      changeId: input.changeId,
      location: 'Sheet!A1',
      inverse: { op: 'restore-cells', range: 'Sheet!A1', values: [[1]] },
      verification: { status: 'unknown', message: 'The write landed, but readback failed.' },
      recoveryPending: true,
      provenanceDropped: true,
    };
    const host = { identity: 'host-instance' };
    const handler = vi.fn(async () => result);
    const dispatcher = createBridgeDispatch('excel', { 'write-cells': handler });
    expect(await dispatcher.dispatch(host, input)).toBe(result);
    expect(handler).toHaveBeenCalledWith(host, input);
  });

  it('preserves host exceptions without claiming that no mutation occurred', async () => {
    const error = new Error('The host disconnected after dispatch');
    const dispatcher = createBridgeDispatch('excel', {
      'write-cells': async () => {
        throw error;
      },
    });
    await expect(dispatcher.dispatch({}, request())).rejects.toBe(error);
  });

  it.each([true, false])(
    'reports unavailable durable provenance for a successful host outcome (payload=%s)',
    async (present) => {
      const input = request();
      if (present)
        input.provenance = {
          agentId: 'test',
          identity: 'user@example.com',
          timestamp: '2026-09-07T00:00:00Z',
          contentHash: 'hash',
          sources: [],
        };
      const dispatcher = createDispatch(
        'excel',
        {
          'write-cells': async (_host, req) => ({
            ok: true,
            kind: req.kind,
            changeId: req.changeId,
          }),
        },
        { provenance: 'unsupported' },
      );
      const result = await dispatcher.dispatch({}, input);
      expect(result).toMatchObject(
        present ? { provenanceDropped: true } : { provenanceMissing: true },
      );
      expect(present ? result.provenanceMissing : result.provenanceDropped).toBeUndefined();
    },
  );

  it('does not label a rejected write as a landed provenance failure', async () => {
    const dispatcher = createDispatch(
      'excel',
      {
        'write-cells': async (_host, req) => ({
          ok: false,
          kind: req.kind,
          changeId: req.changeId,
        }),
      },
      { provenance: 'unsupported' },
    );
    expect(await dispatcher.dispatch({}, request())).toEqual({
      ok: false,
      kind: 'write-cells',
      changeId: 'dispatch-1',
    });
  });
});
