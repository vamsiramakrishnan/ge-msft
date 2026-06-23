import { describe, expect, it } from 'vitest';
import type { DocBridge } from '@ge/runtime';
import { OUTLOOK_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, OutlookBridge } from './outlook-bridge.js';

/**
 * ADR-0006 capability closure — Outlook. Self-contained. `create-mail` was un-advertised (never
 * handled), so advertised==handled=={reply-mail}; no addressable read verbs.
 */
describe('Outlook capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds', () => {
    const advertised = new Set(OUTLOOK_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(HANDLED_ACTUATIONS));
  });

  it('does NOT advertise the previously-phantom create-mail', () => {
    expect(OUTLOOK_CAPABILITIES.actuations.map((a) => a.kind)).not.toContain('create-mail');
  });

  it('declares no read verbs (no addressable read port implemented)', () => {
    expect(OUTLOOK_CAPABILITIES.reads ?? []).toEqual([]);
    const bridge: DocBridge = new OutlookBridge();
    expect(bridge.captureDocState).toBeUndefined();
    expect(bridge.readRange).toBeUndefined();
    expect(bridge.searchDocument).toBeUndefined();
  });
});
