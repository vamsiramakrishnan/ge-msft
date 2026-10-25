import { describe, expect, it, vi } from 'vitest';
import { asChangeId, type ActuationRequest, type CapabilityManifest } from '@ge/contracts';
import { capabilitySupport, hostCapabilityProfile, negotiateCapabilities } from './host-profile.js';
import { createBridgeDispatch } from './bridge-dispatch.js';

const manifest: CapabilityManifest = {
  surface: 'excel',
  contextKinds: [],
  actuations: [
    {
      kind: 'write-cells',
      surface: 'excel',
      title: 'Write',
      reversible: true,
      appliesTo: ['range'],
    },
    {
      kind: 'add-comment',
      surface: 'excel',
      title: 'Comment',
      reversible: true,
      appliesTo: ['range'],
    },
  ],
};
describe('live host negotiation', () => {
  it('checks requirement versions as exact strings and excludes unsupported members', () => {
    const check = vi.fn((_name: string, version: string) => version === '1.1');
    const profile = hostCapabilityProfile('excel', { isSetSupported: check });
    expect(check).toHaveBeenCalledWith('ExcelApi', '1.10');
    expect(negotiateCapabilities(manifest, profile).actuations.map((entry) => entry.kind)).toEqual([
      'write-cells',
    ]);
    expect(manifest.actuations).toHaveLength(2);
  });
  it('fails closed when probing throws or a document is read-only', () => {
    const profile = hostCapabilityProfile('excel', {
      mode: 'read',
      isSetSupported: () => {
        throw new Error('host');
      },
    });
    expect(negotiateCapabilities(manifest, profile).actuations).toEqual([]);
    expect(capabilitySupport('word', 'tracked-change', profile).reason).toBe('wrong_surface');
  });
  it('offers compose edits only on the open draft', () => {
    const profile = hostCapabilityProfile('outlook', { mode: 'read', isSetSupported: () => true });
    expect(capabilitySupport('outlook', 'set-subject', profile).reason).toBe('compose_required');
    expect(capabilitySupport('outlook', 'reply-mail', profile).supported).toBe(true);
  });
  it('rechecks support immediately before dispatch, including after discovery changes', async () => {
    let writable = true;
    const handler = vi.fn(async (_host: object, request: ActuationRequest) => ({
      ...request,
      ok: true,
    }));
    const dispatch = createBridgeDispatch(
      'excel',
      { 'write-cells': handler },
      {
        provenance: 'reported',
        capabilities: () => ({ ...manifest, actuations: writable ? manifest.actuations : [] }),
      },
    );
    writable = false;
    const result = await dispatch.dispatch(
      {},
      {
        surface: 'excel',
        kind: 'write-cells',
        changeId: asChangeId('unavailable'),
        params: { target: { range: 'A1' }, cells: [['1']] },
      },
    );
    expect(result.error?.code).toBe('unsupported');
    expect(handler).not.toHaveBeenCalled();
  });
  it('refuses parameters belonging to another capability before host access', async () => {
    const handler = vi.fn();
    const dispatch = createBridgeDispatch(
      'excel',
      { 'write-cells': handler },
      { provenance: 'reported' },
    );
    const result = await dispatch.dispatch(
      {},
      {
        surface: 'excel',
        kind: 'write-cells',
        changeId: asChangeId('wrong-params'),
        params: { target: { range: 'A1' }, cells: [['1']], mail: { subject: 'wrong' } },
      },
    );
    expect(result.error?.code).toBe('invalid_request');
    expect(handler).not.toHaveBeenCalled();
  });
});
