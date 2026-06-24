import { z } from 'zod';
import { ActuationKindSchema } from './capability.js';
import { READ_VERBS, CONTROL_VERBS, WRITE_VERB_TO_KIND } from './command-grammar.js';
import { TRANSFORM_NAMES, EFFECT_VERBS } from './expr-grammar.js';

/**
 * ADR-0008 — the versioned **language manifest**: the single, serializable description of the
 * `m365-cli` language, emitted from `@ge/contracts` (the one source of truth) so that the skill's
 * Python preflight (`surface-cli`, which runs in a Python/Bash sandbox) can be **generated** from it
 * rather than hand-mirrored. This kills the grammar-drift trap (ADR-0008 §4): the TS grammar, the TS
 * runtime parser, and the generated Python tables all derive from this one object, and the parity
 * corpus proves `generated-python ≡ TS-grammar` on every release.
 *
 * This manifest describes the LANGUAGE (the verbs, value types, pure transforms, effect terminals,
 * and verb→kind map) — NOT the per-turn *capabilities*. Which verbs are live on a given turn still
 * comes from the injected capability signature (`grammarFor`), scoped to a surface's advertised
 * `ActuationKind`s. The language is stable; the per-turn capability slice is dynamic.
 *
 * It deliberately omits prose (usage/hint strings stay in `grammarFor` for the live UI) and policy
 * limits (effect/cell budgets are a separate policy decision; ADR-0008 Phase 2/3). A bump to
 * {@link LANGUAGE_VERSION} is required whenever the verb set, transform set, value types, or
 * verb→kind map changes.
 */
export const LANGUAGE_VERSION = 'm365-cli/1.0' as const;

/**
 * The value algebra (ADR-0008 §1). Pure operators consume and produce these; effect operators
 * consume a value and terminate. `Table/Number/Text/Boolean` are the pipeline value layer;
 * `Selector/RangeRef` are read/anchor inputs. (Deferred: `DocumentRef/SlideRef/MessageRef/
 * ArtifactRef` — added when cross-artifact composition lands; not emitted until the runtime
 * produces them.)
 */
export const VALUE_TYPES = ['Table', 'Number', 'Text', 'Boolean', 'Selector', 'RangeRef'] as const;
export type ValueType = (typeof VALUE_TYPES)[number];

/** The serialized language manifest (the shape bundled as `m365-cli-<v>.json`). */
export const LanguageManifestSchema = z.object({
  version: z.literal(LANGUAGE_VERSION),
  valueTypes: z.array(z.string()),
  verbs: z.object({
    read: z.array(z.string()), // pipeline/command sources (outline/read/search)
    control: z.array(z.string()), // done/help — not actuations
    write: z.array(z.string()), // effect verbs reachable from the model (keys of writeVerbToKind)
  }),
  /** Each write verb → the single `ActuationKind` it compiles to. Many verbs → one kind is allowed. */
  writeVerbToKind: z.record(z.string()),
  /** The pure pipeline transforms (compose freely; never terminate). */
  transforms: z.array(z.string()),
  /** Effect verbs that terminate computation — they can never be a pipeline source. */
  effectVerbs: z.array(z.string()),
  /** The full actuation-kind catalogue (the verb→kind values must be a subset of this). */
  actuationKinds: z.array(z.string()),
});
export type LanguageManifest = z.infer<typeof LanguageManifestSchema>;

/**
 * Build the language manifest from the authoritative `@ge/contracts` definitions. Pure and
 * deterministic (sorted sets) so the emitted JSON is stable across builds — a byte-for-byte diff is
 * a real language change, never incidental ordering. Run through {@link LanguageManifestSchema} by
 * the emitter/consumers; {@link assertManifestConsistent} cross-checks internal invariants.
 */
export function buildLanguageManifest(): LanguageManifest {
  return {
    version: LANGUAGE_VERSION,
    valueTypes: [...VALUE_TYPES],
    verbs: {
      read: [...READ_VERBS],
      control: [...CONTROL_VERBS],
      write: Object.keys(WRITE_VERB_TO_KIND).sort(),
    },
    writeVerbToKind: { ...WRITE_VERB_TO_KIND },
    transforms: [...TRANSFORM_NAMES].sort(),
    effectVerbs: [...EFFECT_VERBS].sort(),
    actuationKinds: [...ActuationKindSchema.options].sort(),
  };
}

/**
 * Assert the manifest's internal invariants — the guard that the emitter itself cannot drift:
 *   1. every write verb maps to a kind in the actuation catalogue (no phantom kinds);
 *   2. every model-reachable write verb is an effect terminal (it appears in `effectVerbs`);
 *   3. transforms and effect verbs are disjoint (a name is pure XOR terminal).
 * Throws an `Error` listing every violation, or returns the validated manifest.
 */
export function assertManifestConsistent(
  manifest: LanguageManifest = buildLanguageManifest(),
): LanguageManifest {
  const parsed = LanguageManifestSchema.parse(manifest);
  const kinds = new Set(parsed.actuationKinds);
  const effects = new Set(parsed.effectVerbs);
  const transforms = new Set(parsed.transforms);
  const errors: string[] = [];

  for (const [verb, kind] of Object.entries(parsed.writeVerbToKind)) {
    if (!kinds.has(kind)) errors.push(`write verb "${verb}" → unknown kind "${kind}"`);
    if (!effects.has(verb)) errors.push(`write verb "${verb}" is not an effect terminal`);
  }
  for (const name of transforms) {
    if (effects.has(name)) errors.push(`"${name}" is both a transform and an effect terminal`);
  }

  if (errors.length > 0) {
    throw new Error(`language manifest is inconsistent:\n  - ${errors.join('\n  - ')}`);
  }
  return parsed;
}
