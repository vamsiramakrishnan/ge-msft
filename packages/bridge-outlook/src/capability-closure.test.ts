import { asChangeId, type ActuationRequest } from '@ge/contracts';
import { describe, expect, it } from 'vitest';
import { implementedRegistryKindsForSurface } from '../../contracts/src/capability-registry.js';
import type { DocBridge } from '@ge/runtime';
import { OUTLOOK_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, OutlookBridge } from './outlook-bridge.js';

/**
 * ADR-0006 capability closure — Outlook. Self-contained. `reply-mail` (reviewable reply form),
 * `create-mail` (new draft via `displayNewMessageForm`), and the four in-place draft edits
 * (`set-recipients`, `add-attachment`, `set-body`, `set-subject`) are advertised AND handled, so
 * advertised==handled; and Outlook declares whole-item `read` + `search`, each backed by a real
 * bridge port (Team READS): `read` → `captureDocState` (a mail item has no addressable sub-range),
 * `search` → `searchDocument`. No `outline` (a mail has no headings).
 */
describe('Outlook capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds', () => {
    const advertised = new Set(OUTLOOK_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(HANDLED_ACTUATIONS));
  });

  it('advertised actuation kinds === registry implemented Outlook capabilities', () => {
    const advertised = new Set(OUTLOOK_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(implementedRegistryKindsForSurface('outlook')));
  });

  it('advertises reply-mail, create-mail, and the four draft edits (each backed by a real actuate case)', () => {
    expect(OUTLOOK_CAPABILITIES.actuations.map((a) => a.kind).sort()).toEqual([
      'add-attachment',
      'create-mail',
      'reply-mail',
      'set-body',
      'set-recipients',
      'set-subject',
    ]);
  });

  it('advertised reads match the implemented read ports', () => {
    expect(new Set(OUTLOOK_CAPABILITIES.reads)).toEqual(new Set(['read', 'search']));
  });

  it('exposes a read port for each advertised read verb (no phantom read)', () => {
    const bridge: DocBridge = new OutlookBridge();
    const reads = new Set(OUTLOOK_CAPABILITIES.reads ?? []);
    // read → whole-item captureDocState (no addressable sub-range, so no readRange);
    // search → searchDocument.
    if (reads.has('read')) expect(typeof bridge.captureDocState).toBe('function');
    if (reads.has('search')) expect(typeof bridge.searchDocument).toBe('function');
  });

  it('does NOT advertise outline (a mail item has no heading structure)', () => {
    expect(OUTLOOK_CAPABILITIES.reads ?? []).not.toContain('outline');
  });
});

describe('OutlookBridge dispatch admission', () => {
  it('rejects another surface before reaching a host handler', async () => {
    const request: ActuationRequest = {
      changeId: asChangeId('wrong-surface'),
      kind: 'reply-mail',
      surface: 'excel',
      params: {},
    };
    expect(await new OutlookBridge().actuate(request)).toMatchObject({
      ok: false,
      error: { code: 'surface_mismatch' },
    });
  });

  it('rejects malformed parameters before reaching a host handler', async () => {
    const request = {
      changeId: asChangeId('invalid-params'),
      kind: 'reply-mail',
      surface: 'outlook',
      params: { text: 42 },
    } as unknown as ActuationRequest;
    expect(await new OutlookBridge().actuate(request)).toMatchObject({
      ok: false,
      error: { code: 'invalid_request' },
    });
  });
});
