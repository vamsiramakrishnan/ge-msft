import { describe, it, expect } from 'vitest';
import { QUICK_ACTIONS, type QuickAction } from '@ge/contracts';
import { quickActionSeed } from './quick-action-seed.js';

function action(over: Partial<QuickAction>): QuickAction {
  return {
    id: 'x',
    label: 'X',
    surfaces: ['word'],
    intent: 'assist',
    prompt: 'do the thing',
    ground: [],
    output: 'chat',
    contextMenu: false,
    ...over,
  };
}

describe('quickActionSeed', () => {
  it('prepends a single @-ground mention ahead of the prompt', () => {
    expect(quickActionSeed(action({ ground: ['this'], prompt: 'Summarize this.' }))).toBe(
      '@this Summarize this.',
    );
  });

  it('joins multiple grounds, in order, before the prompt', () => {
    expect(quickActionSeed(action({ ground: ['this', 'unit'], prompt: 'Review it.' }))).toBe(
      '@this @unit Review it.',
    );
  });

  it('degrades to just the (trimmed) prompt when there is no ground', () => {
    expect(quickActionSeed(action({ ground: [], prompt: '  Catch me up.  ' }))).toBe(
      'Catch me up.',
    );
  });

  it('reads exactly like a composer line for every catalog action (cross-package interplay)', () => {
    for (const a of QUICK_ACTIONS) {
      const seed = quickActionSeed(a);
      expect(seed).toContain(a.prompt.trim());
      if (a.ground.length > 0) {
        expect(seed.startsWith(`@${a.ground[0]}`)).toBe(true);
        // Every ground appears as an @-mention.
        for (const g of a.ground) expect(seed).toContain(`@${g}`);
      }
    }
  });
});
