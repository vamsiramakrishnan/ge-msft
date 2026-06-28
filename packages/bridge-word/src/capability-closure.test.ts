import { describe, expect, it } from 'vitest';
import { WORD_CAPABILITIES } from './capabilities.js';
import { HANDLED_ACTUATIONS, WordBridge } from './word-bridge.js';

/**
 * ADR-0006 capability closure — Word. These assertions are self-contained (no import from the
 * runtime/contracts closure helper) so the bridge stays independently verifiable: the advertised
 * manifest and what `actuate()` actually handles must be the SAME set, and the advertised `reads`
 * must match the bridge's actually-implemented read ports. A reintroduced phantom (an advertised
 * kind with no `actuate()` case, or vice-versa) fails this test and the build.
 */
describe('Word capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds (no phantom, none handled-but-hidden)', () => {
    const advertised = new Set(WORD_CAPABILITIES.actuations.map((a) => a.kind));
    const handled = new Set(HANDLED_ACTUATIONS);
    expect(advertised).toEqual(handled);
  });

  it('advertises the bridge-backed direct Word write kinds', () => {
    const advertised = WORD_CAPABILITIES.actuations.map((a) => a.kind);
    expect(advertised).toEqual([
      'insert-text',
      'replace-selection',
      'insert-ooxml',
      'tracked-change',
      'fill-content-control',
      'add-comment',
      'comment-reply',
    ]);
  });

  it('advertised reads match the implemented read ports', () => {
    // Word serves outline + whole-document read via captureDocState, search via searchDocument.
    expect(new Set(WORD_CAPABILITIES.reads)).toEqual(new Set(['outline', 'read', 'search']));
  });

  it('every advertised read verb has a corresponding bridge read port', () => {
    const bridge = new WordBridge();
    for (const read of WORD_CAPABILITIES.reads ?? []) {
      if (read === 'outline' || read === 'read')
        expect(typeof bridge.captureDocState).toBe('function');
      if (read === 'search') expect(typeof bridge.searchDocument).toBe('function');
    }
  });
});
