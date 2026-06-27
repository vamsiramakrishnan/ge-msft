import { describe, it, expect } from 'vitest';
import { QUICK_ACTIONS, type QuickAction } from '@ge/contracts';
import { quickActionSeed, quickActionToInvocation, invocationToSeed } from './quick-action-seed.js';
import type { ComposerInvocation } from './Composer.js';

function action(over: Partial<QuickAction>): QuickAction {
  return {
    id: 'x',
    label: 'X',
    surfaces: ['word'],
    intent: 'ask',
    scope: { kind: 'selection' },
    prompt: 'do the thing',
    ground: [],
    output: 'chat',
    contextMenu: false,
    ...over,
  };
}

describe('quickActionToInvocation', () => {
  it('carries the typed intent / scope / ground-as-mentions / prompt through (no @-magic-string)', () => {
    const inv = quickActionToInvocation(
      action({
        intent: 'summarize',
        scope: { kind: 'document' },
        ground: ['this'],
        prompt: 'Sum.',
      }),
    );
    expect(inv.intent).toBe('summarize');
    expect(inv.scope).toEqual({ kind: 'document' });
    expect(inv.mentions).toEqual([{ kind: 'this' }]);
    expect(inv.instruction).toBe('Sum.');
  });

  it('substitutes `{{name}}` slots from the collected values (H)', () => {
    const inv = quickActionToInvocation(
      action({ intent: 'draft', prompt: 'Draft on {{topic}}.' }),
      { topic: 'Q3 GTM' },
    );
    expect(inv.instruction).toBe('Draft on Q3 GTM.');
  });

  it('leaves an unprovided slot intact (the fail-closed guard catches it downstream)', () => {
    const inv = quickActionToInvocation(action({ intent: 'draft', prompt: 'Draft on {{topic}}.' }));
    expect(inv.instruction).toBe('Draft on {{topic}}.');
  });
});

describe('invocationToSeed (deterministic string for the controller seam)', () => {
  const base: ComposerInvocation = {
    intent: 'ask',
    scope: { kind: 'selection' },
    mentions: [{ kind: 'this' }],
    instruction: 'Summarize this.',
    raw: '',
  };

  it('renders /verb, mentions, then instruction — like a line a user could have typed', () => {
    expect(invocationToSeed(base)).toBe('/ask @this Summarize this.');
  });

  it('omits the implicit selection scope but emits a scope: token for a non-selection scope', () => {
    expect(invocationToSeed({ ...base, scope: { kind: 'selection' } })).not.toContain('scope:');
    expect(invocationToSeed({ ...base, scope: { kind: 'document' } })).toContain('scope:document');
    expect(invocationToSeed({ ...base, scope: { kind: 'range', ref: 'A1:B2' } })).toContain(
      'scope:range(A1:B2)',
    );
  });

  it('joins multiple grounds, in order', () => {
    expect(invocationToSeed({ ...base, mentions: [{ kind: 'this' }, { kind: 'unit' }] })).toBe(
      '/ask @this @unit Summarize this.',
    );
  });

  it('drops empty parts (no intent → no leading verb)', () => {
    expect(invocationToSeed({ ...base, intent: undefined, mentions: [] })).toBe('Summarize this.');
  });
});

describe('quickActionSeed', () => {
  it('builds the deterministic seed from a catalog action', () => {
    expect(
      quickActionSeed(action({ intent: 'summarize', ground: ['this'], prompt: 'Summarize this.' })),
    ).toBe('/summarize @this Summarize this.');
  });

  it('produces a parseable, grounded seed for every catalog action (cross-package interplay)', () => {
    for (const a of QUICK_ACTIONS) {
      const seed = quickActionSeed(a);
      expect(seed).toContain(a.prompt.trim());
      expect(seed.startsWith(`/${a.intent}`)).toBe(true);
      for (const g of a.ground) expect(seed).toContain(`@${g}`);
    }
  });
});
