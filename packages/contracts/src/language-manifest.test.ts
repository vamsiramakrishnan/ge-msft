import { describe, it, expect } from 'vitest';
import {
  buildLanguageManifest,
  assertManifestConsistent,
  LanguageManifestSchema,
  LANGUAGE_VERSION,
  VALUE_TYPES,
} from './language-manifest.js';
import { READ_VERBS, CONTROL_VERBS, WRITE_VERB_TO_KIND } from './command-grammar.js';
import { TRANSFORM_NAMES, EFFECT_VERBS } from './expr-grammar.js';
import { ActuationKindSchema } from './capability.js';

describe('language manifest (ADR-0008 single source)', () => {
  const manifest = buildLanguageManifest();

  it('validates against its own schema and stamps the version', () => {
    expect(() => LanguageManifestSchema.parse(manifest)).not.toThrow();
    expect(manifest.version).toBe(LANGUAGE_VERSION);
  });

  it('mirrors the authoritative grammar sets exactly (no drift)', () => {
    // The whole point: the manifest IS the TS grammar, serialized — so the generated Python
    // preflight derived from it cannot diverge from the runtime parser.
    expect(manifest.verbs.read).toEqual([...READ_VERBS]);
    expect(manifest.verbs.control).toEqual([...CONTROL_VERBS]);
    expect(manifest.verbs.write).toEqual(Object.keys(WRITE_VERB_TO_KIND).sort());
    expect(manifest.writeVerbToKind).toEqual({ ...WRITE_VERB_TO_KIND });
    expect(manifest.transforms).toEqual([...TRANSFORM_NAMES].sort());
    expect(manifest.effectVerbs).toEqual([...EFFECT_VERBS].sort());
    expect(manifest.actuationKinds).toEqual([...ActuationKindSchema.options].sort());
    expect(manifest.valueTypes).toEqual([...VALUE_TYPES]);
  });

  it('is internally consistent — verbs map to real kinds, write verbs are effect terminals', () => {
    expect(() => assertManifestConsistent(manifest)).not.toThrow();
  });

  it('every write verb compiles to a kind in the actuation catalogue', () => {
    const kinds = new Set(manifest.actuationKinds);
    for (const kind of Object.values(manifest.writeVerbToKind)) {
      expect(kinds.has(kind)).toBe(true);
    }
  });

  it('transforms and effect terminals are disjoint (pure XOR terminal)', () => {
    const effects = new Set(manifest.effectVerbs);
    for (const t of manifest.transforms) expect(effects.has(t)).toBe(false);
  });

  it('is deterministic — two builds are byte-identical (stable emitted JSON)', () => {
    expect(JSON.stringify(buildLanguageManifest())).toBe(JSON.stringify(buildLanguageManifest()));
  });

  it('rejects an inconsistent manifest (phantom kind) with a listed error', () => {
    const broken = {
      ...manifest,
      writeVerbToKind: { ...manifest.writeVerbToKind, set: 'no-such-kind' },
    };
    expect(() => assertManifestConsistent(broken)).toThrow(/unknown kind "no-such-kind"/);
  });
});
