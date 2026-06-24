import { describe, it, expect } from 'vitest';
import { IntentSchema, type Intent } from './intent.js';
import { commandPaletteFor } from './command-palette.js';

const SURFACES = ['word', 'excel', 'powerpoint', 'onenote', 'outlook', 'teams'] as const;

describe('commandPaletteFor', () => {
  it('returns a non-empty palette that always offers /assist on every surface', () => {
    for (const surface of SURFACES) {
      const spec = commandPaletteFor(surface);
      expect(spec.surface).toBe(surface);
      expect(spec.verbs.length).toBeGreaterThan(0);
      expect(spec.verbs.some((v) => v.intent === 'assist')).toBe(true);
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

  it('covers every Intent with a verb mapping (so no surface can reference a missing verb)', () => {
    // The union of all palettes must only use intents that have a defined verb; conversely every
    // intent that any surface offers must round-trip to a label.
    const offered = new Set<Intent>(
      SURFACES.flatMap((s) => commandPaletteFor(s).verbs.map((v) => v.intent)),
    );
    for (const intent of offered) {
      const spec = SURFACES.map((s) => commandPaletteFor(s).verbs)
        .flat()
        .find((v) => v.intent === intent);
      expect(spec?.label).toMatch(/^\//);
    }
  });

  it('offers the fixed five mention kinds, as a fresh array each call', () => {
    const a = commandPaletteFor('word');
    const b = commandPaletteFor('word');
    expect(a.mentionKinds).toEqual(['document', 'person', 'datastore', 'this', 'upload']);
    // Mutating one palette's array must not bleed into another (no shared module-level reference).
    a.mentionKinds.push('upload');
    expect(b.mentionKinds).toEqual(['document', 'person', 'datastore', 'this', 'upload']);
  });

  it('gives Word the in-document verbs and Excel a narrower set', () => {
    const word = commandPaletteFor('word').verbs.map((v) => v.label);
    expect(word).toEqual(expect.arrayContaining(['/assist', '/review', '/resolve', '/rewrite']));
    const excel = commandPaletteFor('excel').verbs.map((v) => v.label);
    expect(excel).not.toContain('/rewrite'); // Word-only
  });

  it('narrows the palette by allowed intents (capability closure, ADR-0006)', () => {
    const assistOnly = commandPaletteFor('word', ['assist']);
    expect(assistOnly.verbs.map((v) => v.intent)).toEqual(['assist']);
    // An empty closure yields no verbs (but mention kinds remain).
    const none = commandPaletteFor('word', []);
    expect(none.verbs).toEqual([]);
    expect(none.mentionKinds.length).toBe(5);
  });

  it('accepts a Set of allowed intents and never offers a verb outside the surface set', () => {
    const spec = commandPaletteFor('excel', new Set<Intent>(['assist', 'review', 'draft-slides']));
    // draft-slides is allowed but NOT in Excel's surface set → still absent.
    expect(spec.verbs.map((v) => v.intent)).toEqual(['assist', 'review']);
  });
});
