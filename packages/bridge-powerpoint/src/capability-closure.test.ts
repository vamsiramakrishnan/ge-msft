import { describe, expect, it } from 'vitest';
import type { DocBridge } from '@ge/runtime';
import { POWERPOINT_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, PowerPointBridge } from './powerpoint-bridge.js';

/**
 * ADR-0006 capability closure — PowerPoint. Self-contained. `set-speaker-notes` was un-advertised
 * (it always degraded), so advertised==handled=={insert-slide}; and PowerPoint now declares the
 * `outline`/`read`/`search` read verbs, each backed by a real bridge port (Team READS / ADR-0006).
 */
describe('PowerPoint capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds', () => {
    const advertised = new Set(POWERPOINT_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(HANDLED_ACTUATIONS));
  });

  it('does NOT advertise the previously-phantom set-speaker-notes', () => {
    expect(POWERPOINT_CAPABILITIES.actuations.map((a) => a.kind)).not.toContain(
      'set-speaker-notes',
    );
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
