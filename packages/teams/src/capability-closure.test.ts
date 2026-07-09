import { describe, expect, it } from 'vitest';
import { implementedRegistryKindsForSurface } from '../../contracts/src/capability-registry.js';
import type { DocBridge } from '@ge/runtime';
import { TEAMS_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, TeamsBridge } from './teams-bridge.js';

/**
 * ADR-0006 capability closure — Teams. Self-contained. advertised==handled=={post-message}; and
 * Teams now declares whole-transcript `read` + `search`, each backed by a real bridge port (Team
 * READS): `read` → `captureDocState` (a transcript has no addressable sub-range), `search` →
 * `searchDocument`. No `outline` (a transcript has no heading structure).
 */
describe('Teams capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds', () => {
    const advertised = new Set(TEAMS_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(HANDLED_ACTUATIONS));
  });

  it('advertised actuation kinds === registry implemented Teams capabilities', () => {
    const advertised = new Set(TEAMS_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(implementedRegistryKindsForSurface('teams')));
  });

  it('advertised reads match the implemented read ports', () => {
    expect(new Set(TEAMS_CAPABILITIES.reads)).toEqual(new Set(['read', 'search']));
  });

  it('exposes a read port for each advertised read verb (no phantom read)', () => {
    const bridge: DocBridge = new TeamsBridge();
    const reads = new Set(TEAMS_CAPABILITIES.reads ?? []);
    // read → whole-transcript captureDocState (no addressable sub-range, so no readRange);
    // search → searchDocument.
    if (reads.has('read')) expect(typeof bridge.captureDocState).toBe('function');
    if (reads.has('search')) expect(typeof bridge.searchDocument).toBe('function');
  });

  it('does NOT advertise outline (a transcript has no heading structure)', () => {
    expect(TEAMS_CAPABILITIES.reads ?? []).not.toContain('outline');
  });
});
