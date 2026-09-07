import { asChangeId, type ActuationRequest } from '@ge/contracts';
import { describe, expect, it } from 'vitest';
import { implementedRegistryKindsForSurface } from '../../contracts/src/capability-registry.js';
import type { DocBridge } from '@ge/runtime';
import { POWERPOINT_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, PowerPointBridge } from './powerpoint-bridge.js';

/**
 * ADR-0006 capability closure — PowerPoint. Self-contained. `set-speaker-notes` was un-advertised
 * (it always degraded), and `insert-image` is likewise un-advertised (no image-insertion write
 * path in the pinned Office.js typings), so advertised==handled=={insert-slide, set-shape-text,
 * add-shape, format-shape, add-table-slide}; and PowerPoint declares the `outline`/`read`/`search`
 * read verbs, each backed by a real bridge port (Team READS / ADR-0006).
 */
describe('PowerPoint capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds', () => {
    const advertised = new Set(POWERPOINT_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(HANDLED_ACTUATIONS));
  });

  it('advertised actuation kinds === registry implemented PowerPoint capabilities', () => {
    const advertised = new Set(POWERPOINT_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(implementedRegistryKindsForSurface('powerpoint')));
  });

  it('does NOT advertise the previously-phantom set-speaker-notes', () => {
    expect(POWERPOINT_CAPABILITIES.actuations.map((a) => a.kind)).not.toContain(
      'set-speaker-notes',
    );
  });

  it('does NOT advertise insert-image (no typed host write path in the pinned typings)', () => {
    expect(POWERPOINT_CAPABILITIES.actuations.map((a) => a.kind)).not.toContain('insert-image');
  });

  it('advertised reads match the implemented read ports', () => {
    expect(new Set(POWERPOINT_CAPABILITIES.reads)).toEqual(new Set(['outline', 'read', 'search']));
  });

  it('exposes a read port for each advertised read verb (no phantom read)', () => {
    const bridge: DocBridge = new PowerPointBridge();
    const reads = new Set(POWERPOINT_CAPABILITIES.reads ?? []);
    // outline → captureDocState; read → readRange (addressable slide); search → searchDocument.
    if (reads.has('outline')) expect(typeof bridge.captureDocState).toBe('function');
    if (reads.has('read')) expect(typeof bridge.readRange).toBe('function');
    if (reads.has('search')) expect(typeof bridge.searchDocument).toBe('function');
  });
});

describe('PowerPointBridge dispatch admission', () => {
  it('rejects another surface before reaching a host handler', async () => {
    const request: ActuationRequest = {
      changeId: asChangeId('wrong-surface'),
      kind: 'insert-slide',
      surface: 'excel',
      params: {},
    };
    expect(await new PowerPointBridge().actuate(request)).toMatchObject({
      ok: false,
      error: { code: 'surface_mismatch' },
    });
  });

  it('rejects malformed parameters before reaching a host handler', async () => {
    const request = {
      changeId: asChangeId('invalid-params'),
      kind: 'insert-slide',
      surface: 'powerpoint',
      params: { text: 42 },
    } as unknown as ActuationRequest;
    expect(await new PowerPointBridge().actuate(request)).toMatchObject({
      ok: false,
      error: { code: 'invalid_request' },
    });
  });
});
