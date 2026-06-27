import { describe, it, expect } from 'vitest';
import { IntentSchema, type Intent } from './intent.js';
import { commandPaletteFor, VERBS_BY_SURFACE } from './command-palette.js';

const SURFACES = ['word', 'excel', 'powerpoint', 'onenote', 'outlook', 'teams'] as const;

describe('commandPaletteFor', () => {
  it('returns a non-empty palette that always offers /ask on every surface', () => {
    for (const surface of SURFACES) {
      const spec = commandPaletteFor(surface);
      expect(spec.surface).toBe(surface);
      expect(spec.verbs.length).toBeGreaterThan(0);
      expect(spec.verbs.some((v) => v.intent === 'ask')).toBe(true);
    }
  });

  it('maps each verb to a defined /label and a description (no missing VERB_BY_INTENT entry)', () => {
    for (const surface of SURFACES) {
      for (const verb of commandPaletteFor(surface).verbs) {
        expect(verb).toBeDefined();
        expect(verb.label.startsWith('/')).toBe(true);
        expect(verb.description.length).toBeGreaterThan(0);
        // Every offered intent is a real IntentSchema value (no phantom verb).
        expect(IntentSchema.options).toContain(verb.intent);
      }
    }
  });

  it('labels every verb as /<intent> (default rule, no override map)', () => {
    for (const surface of SURFACES) {
      for (const verb of commandPaletteFor(surface).verbs) {
        expect(verb.label).toBe(`/${verb.intent}`);
      }
    }
  });

  it('covers every Intent with a verb mapping (invariant c)', () => {
    // Every IntentSchema value must round-trip to a /label on at least one surface, so no surface
    // can reference a verb the palette cannot render.
    const offered = new Map<Intent, string>();
    for (const surface of SURFACES) {
      for (const v of commandPaletteFor(surface).verbs) offered.set(v.intent, v.label);
    }
    for (const intent of IntentSchema.options) {
      expect(offered.get(intent)).toBe(`/${intent}`);
    }
  });

  it('offers only immediately resolvable @-mention kinds, as a fresh array each call', () => {
    const a = commandPaletteFor('word');
    const b = commandPaletteFor('word');
    expect(a.mentionKinds).toEqual(['this', 'unit']);
    // Mutating one palette's array must not bleed into another (no shared module-level reference).
    a.mentionKinds.push('upload');
    expect(b.mentionKinds).toEqual(['this', 'unit']);
  });

  it('exposes surface-named scope options as data (default is the first entry)', () => {
    const word = commandPaletteFor('word');
    expect(word.scopeOptions.map((o) => o.label)).toEqual([
      'Selection',
      'Whole document',
      'This section',
    ]);
    expect(word.scopeOptions[0]!.scope).toEqual({ kind: 'selection' });
    const excel = commandPaletteFor('excel').scopeOptions.map((o) => o.label);
    expect(excel).toEqual(['Selection', 'Sheet', 'Range']);
    const teams = commandPaletteFor('teams').scopeOptions;
    expect(teams.map((o) => o.label)).toEqual(['Transcript', 'Last 5 min']);
    expect(teams[1]!.scope).toEqual({ kind: 'range', ref: 'last-5-min' });
  });

  it('gives Word the in-document verbs and keeps /draft off Word', () => {
    const word = commandPaletteFor('word').verbs.map((v) => v.label);
    expect(word).toEqual(
      expect.arrayContaining(['/ask', '/summarize', '/explain', '/rewrite', '/review']),
    );
    expect(word).not.toContain('/draft'); // generation lives on PPT/OneNote/Outlook
    const teams = commandPaletteFor('teams').verbs.map((v) => v.label);
    expect(teams).toEqual(expect.arrayContaining(['/ask', '/summarize', '/notes']));
    expect(teams).not.toContain('/explain'); // not in Teams' verb set
  });

  it('offers /visualize on Excel only', () => {
    const excel = commandPaletteFor('excel').verbs.map((v) => v.label);
    expect(excel).toContain('/visualize');
    for (const surface of SURFACES.filter((s) => s !== 'excel')) {
      expect(commandPaletteFor(surface).verbs.map((v) => v.label)).not.toContain('/visualize');
    }
  });

  it('narrows the palette by allowed intents (capability closure, ADR-0006)', () => {
    const askOnly = commandPaletteFor('word', ['ask']);
    expect(askOnly.verbs.map((v) => v.intent)).toEqual(['ask']);
    // An empty closure yields no verbs (but mention kinds + scope options remain).
    const none = commandPaletteFor('word', []);
    expect(none.verbs).toEqual([]);
    expect(none.mentionKinds).toEqual(['this', 'unit']);
    expect(none.scopeOptions.length).toBeGreaterThan(0);
  });

  it('accepts a Set of allowed intents and never offers a verb outside the surface set', () => {
    const spec = commandPaletteFor('excel', new Set<Intent>(['ask', 'review', 'draft']));
    // draft is allowed but NOT in Excel's surface set → still absent.
    expect(spec.verbs.map((v) => v.intent)).toEqual(['ask', 'review']);
  });

  it('every surface verb set is a subset of IntentSchema', () => {
    for (const surface of SURFACES) {
      for (const intent of VERBS_BY_SURFACE[surface]) {
        expect(IntentSchema.options).toContain(intent);
      }
    }
  });
});
