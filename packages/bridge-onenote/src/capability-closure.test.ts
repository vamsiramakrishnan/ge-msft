import { describe, expect, it } from 'vitest';
import { implementedRegistryKindsForSurface } from '../../contracts/src/capability-registry.js';
import type { DocBridge } from '@ge/runtime';
import { ONENOTE_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, OneNoteBridge } from './onenote-bridge.js';

/**
 * ADR-0006 capability closure — OneNote. Self-contained. advertised==handled=={append-page}; and
 * OneNote now declares `outline`/`read`/`search`, each backed by a real bridge port (Team READS):
 * `outline`/whole-page `read` → `captureDocState`, `search` → `searchDocument`.
 */
describe('OneNote capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds', () => {
    const advertised = new Set(ONENOTE_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(HANDLED_ACTUATIONS));
  });

  it('advertised actuation kinds === registry implemented OneNote capabilities', () => {
    const advertised = new Set(ONENOTE_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(implementedRegistryKindsForSurface('onenote')));
  });

  it('advertised reads match the implemented read ports', () => {
    expect(new Set(ONENOTE_CAPABILITIES.reads)).toEqual(new Set(['outline', 'read', 'search']));
  });

  it('exposes a read port for each advertised read verb (no phantom read)', () => {
    const bridge: DocBridge = new OneNoteBridge();
    const reads = new Set(ONENOTE_CAPABILITIES.reads ?? []);
    // outline → captureDocState; read → whole-page captureDocState (no addressable sub-range, so
    // no readRange); search → searchDocument.
    if (reads.has('outline')) expect(typeof bridge.captureDocState).toBe('function');
    if (reads.has('read')) expect(typeof bridge.captureDocState).toBe('function');
    if (reads.has('search')) expect(typeof bridge.searchDocument).toBe('function');
  });
});
