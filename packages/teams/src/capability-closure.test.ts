import { describe, expect, it } from 'vitest';
import type { DocBridge } from '@ge/runtime';
import { TEAMS_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, TeamsBridge } from './teams-bridge.js';

/**
 * ADR-0006 capability closure — Teams. Self-contained. advertised==handled=={post-message};
 * no addressable read verbs (transcript is the universal context port, not a read verb).
 */
describe('Teams capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds', () => {
    const advertised = new Set(TEAMS_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(HANDLED_ACTUATIONS));
  });

  it('declares no read verbs (no addressable read port implemented)', () => {
    expect(TEAMS_CAPABILITIES.reads ?? []).toEqual([]);
    const bridge: DocBridge = new TeamsBridge();
    expect(bridge.captureDocState).toBeUndefined();
    expect(bridge.readRange).toBeUndefined();
    expect(bridge.searchDocument).toBeUndefined();
  });
});
