import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, it, expect } from 'vitest';
import {
  buildLanguageManifest,
  assertManifestConsistent,
  LanguageManifestSchema,
  LANGUAGE_VERSION,
  VALUE_TYPES,
} from './language-manifest.js';
import {
  READ_VERBS,
  WORKSPACE_VERBS,
  CONTROL_VERBS,
  WRITE_VERB_TO_KIND,
} from './command-grammar.js';
import { COMMAND_HELP } from './command-help.js';
import { TRANSFORM_NAMES, EFFECT_VERBS } from './expr-grammar.js';
import { ActuationKindSchema } from './capability.js';
import {
  assertCapabilityRegistryConsistent,
  capabilityRegistryEntries,
} from './capability-registry.js';

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
    expect(manifest.verbs.workspace).toEqual([...WORKSPACE_VERBS]);
    expect(manifest.verbs.control).toEqual([...CONTROL_VERBS]);
    expect(manifest.verbs.write).toEqual(Object.keys(WRITE_VERB_TO_KIND).sort());
    expect(manifest.writeVerbToKind).toEqual({ ...WRITE_VERB_TO_KIND });
    expect(manifest.transforms).toEqual([...TRANSFORM_NAMES].sort());
    expect(manifest.effectVerbs).toEqual([...EFFECT_VERBS].sort());
    expect(manifest.actuationKinds).toEqual([...ActuationKindSchema.options].sort());
    expect(manifest.valueTypes).toEqual([...VALUE_TYPES]);
    expect(manifest.commandHelp).toEqual(
      Object.fromEntries(Object.entries(COMMAND_HELP).sort(([a], [b]) => a.localeCompare(b))),
    );
    expect(manifest.capabilityRegistry).toEqual(
      capabilityRegistryEntries().sort((a, b) =>
        `${a.surface}:${a.kind}`.localeCompare(`${b.surface}:${b.kind}`),
      ),
    );
  });

  it('has a topic-help entry for every emitted verb and no non-language stray topics', () => {
    const verbs = new Set([
      ...manifest.verbs.read,
      ...manifest.verbs.workspace,
      ...manifest.verbs.control,
      ...manifest.verbs.write,
    ]);
    const kinds = new Set(manifest.actuationKinds);
    for (const verb of verbs) expect(manifest.commandHelp[verb]).toBeDefined();
    for (const topic of Object.keys(manifest.commandHelp)) {
      expect(verbs.has(topic) || kinds.has(topic)).toBe(true);
    }
  });

  it('is internally consistent — verbs map to real kinds, write verbs are effect terminals', () => {
    expect(() => assertManifestConsistent(manifest)).not.toThrow();
    expect(() => assertCapabilityRegistryConsistent(manifest.capabilityRegistry)).not.toThrow();
  });

  it('emits advanced capability registry metadata for skill-side progressive disclosure', () => {
    expect(manifest.capabilityRegistry).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          surface: 'excel',
          kind: 'insert-pivot',
          command: '/insert-pivot',
          status: 'promotable',
        }),
        expect.objectContaining({
          surface: 'powerpoint',
          kind: 'format-shape',
          command: '/format-shape',
          status: 'implemented',
        }),
      ]),
    );
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

  it('the committed JSON the skill bundles matches the emitter (drift gate)', () => {
    // ADR-0008 §4: the skill's Python preflight loads this committed file. If the grammar changes
    // without re-running `bun run emit:language`, this fails — the committed artifact can never drift
    // from the TS source of truth.
    const jsonPath = fileURLToPath(
      new URL('../../../skill/m365-surface-commander/scripts/m365-cli-1.0.json', import.meta.url),
    );
    const committed = JSON.parse(readFileSync(jsonPath, 'utf8'));
    expect(committed).toEqual(manifest);
  });
});
