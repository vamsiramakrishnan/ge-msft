import { describe, expect, it } from 'vitest';
import type { DocBridge } from '@ge/runtime';
import { POWERPOINT_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, PowerPointBridge } from './powerpoint-bridge.js';

/**
 * ADR-0006 capability closure — PowerPoint. Self-contained. `set-speaker-notes` was un-advertised
 * (it always degraded), so advertised==handled=={insert-slide}; and PowerPoint declares NO read
 * verbs because it has no captureDocState/readRange/searchDocument port.
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

  it('declares no read verbs (no addressable read port implemented)', () => {
    expect(POWERPOINT_CAPABILITIES.reads ?? []).toEqual([]);
    const bridge: DocBridge = new PowerPointBridge();
    expect(bridge.captureDocState).toBeUndefined();
    expect(bridge.readRange).toBeUndefined();
    expect(bridge.searchDocument).toBeUndefined();
  });
});
