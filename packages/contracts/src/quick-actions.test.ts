import { describe, it, expect } from 'vitest';
import type { Intent } from './intent.js';
import {
  QUICK_ACTIONS,
  UNIVERSAL_ACTIONS,
  SURFACE_ACTIONS,
  CONTRACT_REVIEW_PACK,
  QuickActionSchema,
  deriveOutput,
  quickActionsForSurface,
  promptPlaceholders,
  actionParameters,
  fillPrompt,
  hasUnfilledPlaceholder,
} from './quick-actions.js';
import { INTENT_REQUIRES } from './intent-capability.js';
import { VERBS_BY_SURFACE } from './command-palette.js';

describe('QUICK_ACTIONS catalog', () => {
  it('every quick action parses the schema', () => {
    for (const action of QUICK_ACTIONS) {
      expect(() => QuickActionSchema.parse(action)).not.toThrow();
    }
  });

  it('every action lists at least one surface and a non-empty prompt', () => {
    for (const action of QUICK_ACTIONS) {
      expect(action.surfaces.length).toBeGreaterThan(0);
      expect(action.prompt.length).toBeGreaterThan(0);
    }
  });

  it('ids are unique (universal + surface)', () => {
    const ids = QUICK_ACTIONS.map((a) => a.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('is exactly the universal block followed by the surface block', () => {
    expect(QUICK_ACTIONS).toEqual([...UNIVERSAL_ACTIONS, ...SURFACE_ACTIONS]);
  });
});

describe('invariant: output is derived from intent', () => {
  // (a) output==='write'||'annotation' ⇒ INTENT_REQUIRES[action.intent].length > 0
  it('every write/annotation action has a non-empty INTENT_REQUIRES; every chat action is empty', () => {
    for (const action of [...QUICK_ACTIONS, ...CONTRACT_REVIEW_PACK]) {
      if (action.output === 'write' || action.output === 'annotation') {
        expect(INTENT_REQUIRES[action.intent].length).toBeGreaterThan(0);
      } else {
        expect(action.output).toBe('chat');
        expect(INTENT_REQUIRES[action.intent]).toEqual([]);
      }
    }
  });

  it('every action.output equals deriveOutput(action.intent)', () => {
    for (const action of [...QUICK_ACTIONS, ...CONTRACT_REVIEW_PACK]) {
      expect(action.output).toBe(deriveOutput(action.intent));
    }
  });
});

describe('invariant: closure — every action.intent is offered on each of its surfaces', () => {
  // (b) every QuickAction's intent ∈ VERBS_BY_SURFACE[surface] for each surface it lists
  it('holds for the whole catalog and the vertical pack', () => {
    for (const action of [...QUICK_ACTIONS, ...CONTRACT_REVIEW_PACK]) {
      for (const surface of action.surfaces) {
        expect(VERBS_BY_SURFACE[surface]).toContain(action.intent);
      }
    }
  });
});

describe('re-tagged actions (the audit fixes)', () => {
  const byId = (id: string) =>
    [...QUICK_ACTIONS, ...CONTRACT_REVIEW_PACK].find((a) => a.id === id)!;

  it('draft-reply is a draft (write), not assist', () => {
    expect(byId('draft-reply').intent).toBe('draft');
    expect(byId('draft-reply').output).toBe('write');
  });

  it('risk-column and write-formula are rewrites (write)', () => {
    expect(byId('risk-column').intent).toBe('rewrite');
    expect(byId('write-formula').intent).toBe('rewrite');
    expect(byId('risk-column').output).toBe('write');
  });

  it('create-chart is visualize (write), not chat', () => {
    expect(byId('create-chart').intent).toBe('visualize');
    expect(byId('create-chart').output).toBe('write');
  });

  it('surface write affordances expose existing bridge-backed actions', () => {
    expect(byId('comment-on-issues').intent).toBe('review');
    expect(byId('comment-on-issues').output).toBe('annotation');
    expect(byId('draft-slide').intent).toBe('draft');
    expect(byId('draft-slide').output).toBe('write');
    expect(byId('draft-new-email').intent).toBe('draft');
    expect(byId('draft-new-email').output).toBe('write');
  });

  it('PPT speaker-notes is a chat draft (no host write path)', () => {
    expect(byId('speaker-notes').output).toBe('chat');
  });

  it('every prompt slot is a declared parameter and vice versa (template↔param parity)', () => {
    for (const action of QUICK_ACTIONS) {
      const slots = new Set(promptPlaceholders(action.prompt));
      const params = new Set(actionParameters(action).map((p) => p.name));
      expect([...slots].sort()).toEqual([...params].sort());
    }
  });

  it('the parameterized actions declare their fill slots', () => {
    expect(actionParameters(byId('write-formula')).map((p) => p.name)).toEqual(['goal']);
    expect(actionParameters(byId('draft-section')).map((p) => p.name)).toEqual(['topic']);
    expect(actionParameters(byId('draft-slide')).map((p) => p.name)).toEqual(['topic']);
    expect(actionParameters(byId('draft-new-email')).map((p) => p.name)).toEqual(['subject']);
    expect(actionParameters(byId('ge-ask')).map((p) => p.name)).toEqual(['question', 'range']);
    expect(actionParameters(byId('review-against')).map((p) => p.name)).toEqual(['standard']);
  });

  it('rejects an action whose prompt slot has no declared parameter', () => {
    expect(() =>
      QuickActionSchema.parse({
        id: 'leaky',
        label: 'Leaky',
        surfaces: ['word'],
        intent: 'rewrite',
        scope: { kind: 'selection' },
        prompt: 'Rewrite for {{tone}}.', // {{tone}} undeclared
        ground: [],
        output: 'write',
        contextMenu: false,
      }),
    ).toThrow(/no declared parameter/);
  });

  it('rejects an action that declares a parameter it never references', () => {
    expect(() =>
      QuickActionSchema.parse({
        id: 'dangling',
        label: 'Dangling',
        surfaces: ['word'],
        intent: 'rewrite',
        scope: { kind: 'selection' },
        prompt: 'Rewrite this.', // no {{tone}} slot
        ground: [],
        output: 'write',
        contextMenu: false,
        parameters: [{ name: 'tone', label: 'Tone' }],
      }),
    ).toThrow(/not referenced/);
  });

  it('fillPrompt substitutes provided slots and leaves the rest intact', () => {
    expect(fillPrompt('Draft on {{topic}}.', { topic: 'GTM' })).toBe('Draft on GTM.');
    expect(
      fillPrompt('=GE.ASK("{{question}}", {{range}})', { question: 'growth?', range: 'A1:B2' }),
    ).toBe('=GE.ASK("growth?", A1:B2)');
    expect(fillPrompt('Draft on {{topic}}.', {})).toBe('Draft on {{topic}}.'); // unprovided left intact
  });

  it('hasUnfilledPlaceholder flags a leftover slot (the fail-closed dispatch guard)', () => {
    expect(hasUnfilledPlaceholder('Draft on {{topic}}.')).toBe(true);
    expect(hasUnfilledPlaceholder('Draft on GTM.')).toBe(false);
  });

  it('exposes the general Review against… and a =GE.ASK Excel action', () => {
    expect(byId('review-against').intent).toBe('review');
    expect(byId('review-against').prompt).toContain('{{standard}}');
    expect(byId('ge-ask').intent).toBe('ask');
    expect(byId('ge-ask').prompt).toContain('GE.ASK');
  });

  it('adds the OneNote "add sources to the unit" composition action', () => {
    expect(byId('add-sources-to-unit').surfaces).toContain('onenote');
  });

  it('keeps the contract-review nouns out of the default catalog', () => {
    const ids = QUICK_ACTIONS.map((a) => a.id);
    expect(ids).not.toContain('review-policy');
    expect(ids).not.toContain('find-unsupported');
    expect(CONTRACT_REVIEW_PACK.map((a) => a.id)).toEqual(['review-policy', 'find-unsupported']);
  });
});

describe('QuickActionSchema enforces output === deriveOutput(intent) (security review, Finding 1)', () => {
  const base = {
    id: 'x',
    label: 'X',
    surfaces: ['word'],
    intent: 'ask',
    scope: { kind: 'selection' },
    prompt: 'do it',
    ground: [],
    output: 'chat',
    contextMenu: false,
  };

  it('accepts a matching tuple (chat verb / write verb)', () => {
    expect(() => QuickActionSchema.parse(base)).not.toThrow();
    expect(() =>
      QuickActionSchema.parse({ ...base, intent: 'rewrite', output: 'write' }),
    ).not.toThrow();
  });

  it('rejects a write-labeled chat verb — the routing safety trap', () => {
    // {intent:'ask', output:'write'} would advertise a write but route to send — refuse at parse.
    expect(() => QuickActionSchema.parse({ ...base, intent: 'ask', output: 'write' })).toThrow();
    expect(() => QuickActionSchema.parse({ ...base, intent: 'rewrite', output: 'chat' })).toThrow();
  });

  it('every shipped catalog + pack action satisfies the refine', () => {
    for (const a of [...QUICK_ACTIONS, ...CONTRACT_REVIEW_PACK]) {
      expect(a.output).toBe(deriveOutput(a.intent));
    }
  });
});

describe('quickActionsForSurface', () => {
  it('filters to actions offered on the surface', () => {
    const word = quickActionsForSurface('word');
    expect(word.length).toBeGreaterThan(0);
    expect(word.every((a) => a.surfaces.includes('word'))).toBe(true);
    // A Word action and a universal action are both present; an Excel-only one is not.
    expect(word.some((a) => a.id === 'review-against')).toBe(true);
    expect(word.some((a) => a.id === 'summarize-this')).toBe(true);
    expect(word.some((a) => a.id === 'summarize-range')).toBe(false);
  });

  it('summarize-this (a universal action) appears on every surface', () => {
    const surfaces = ['word', 'excel', 'powerpoint', 'onenote', 'outlook', 'teams'] as const;
    for (const surface of surfaces) {
      const ids = quickActionsForSurface(surface).map((a) => a.id);
      expect(ids).toContain('summarize-this');
    }
  });

  it('drops actions whose intent is not in the allowed set (closure, ADR-0006)', () => {
    // Word offers a `review` action (review-against) — exclude `review` and it disappears.
    const allowed: Intent[] = ['ask', 'summarize', 'explain', 'rewrite'];
    const ids = quickActionsForSurface('word', allowed).map((a) => a.id);
    expect(ids).not.toContain('review-against');
    expect(ids).toContain('tighten'); // rewrite is allowed
    expect(ids).toContain('summarize-this'); // summarize is allowed
  });

  it('accepts a Set of allowed intents', () => {
    const ids = quickActionsForSurface('excel', new Set<Intent>(['summarize'])).map((a) => a.id);
    expect(ids).toContain('summarize-range');
    expect(ids).not.toContain('find-anomalies'); // review excluded
  });
});
