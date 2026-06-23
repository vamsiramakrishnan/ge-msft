import { describe, expect, it } from 'vitest';
import type { DocBridge } from '@ge/runtime';
import { ONENOTE_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, OneNoteBridge } from './onenote-bridge.js';

/**
 * ADR-0006 capability closure — OneNote. Self-contained. advertised==handled=={append-page};
 * no read verbs (no captureDocState/readRange/searchDocument port).
 */
describe('OneNote capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds', () => {
    const advertised = new Set(ONENOTE_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(HANDLED_ACTUATIONS));
  });

  it('declares no read verbs (no addressable read port implemented)', () => {
    expect(ONENOTE_CAPABILITIES.reads ?? []).toEqual([]);
    const bridge: DocBridge = new OneNoteBridge();
    expect(bridge.captureDocState).toBeUndefined();
    expect(bridge.readRange).toBeUndefined();
    expect(bridge.searchDocument).toBeUndefined();
  });
});
