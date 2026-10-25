import {
  COMMAND_HELP,
  grammarFor,
  registryEntryForKindAndSurface,
  WRITE_VERB_TO_KIND,
  type ActuationKind,
  type CapabilityManifest,
  type CommandHelpEntry,
  type VerbSpec,
} from '@ge/contracts';

/** A bounded, capability-scoped command description. Metadata never grants execution authority. */
export interface CommandCard {
  command: string;
  syntax: string;
  useWhen: string;
  prerequisites: string[];
  limits: string[];
  example: string;
}

export const COMMAND_DISCOVERY_LIMIT = 4;

/** Ranking bonus when every word of a command's name appears in the task (see `discoverCommands`). */
const WHOLE_NAME_BONUS = 25;

/**
 * Deterministic lexical discovery over the same grammar/help used by execution. No model, network,
 * document content or dynamic registrations are involved. An unrelated query returns no cards;
 * the query itself is never reflected into the instruction channel.
 */
export function discoverCommands(manifest: CapabilityManifest, query: string): CommandCard[] {
  const specs = grammarFor(manifest);
  // "slide 3" / "slides 2-4" name WHERE to act, not the `slide` command: ranking on them made every
  // "… on slide 1" request surface `slide` instead of the shape/format command it needed.
  const raw = query.slice(0, 1024);
  const located = raw.replace(/\bslides?\s+\d+(?:\s*(?:-|–|to|and|,)\s*\d+)*/gi, ' ');
  const terms = words(located).slice(0, 32);
  if (terms.length === 0) return [];
  const termSet = new Set(terms);
  // The whole-name check may still see the location words ("add a table to slide 3" names all of
  // add-table-slide), but only once some name word was requested outside them.
  const allTerms = new Set(words(raw));
  const allowed = new Set(specs.map((spec) => spec.verb));
  return specs
    .map((spec, order) => {
      const entry = helpFor(manifest, spec);
      const nameWords = words(spec.verb);
      const name = new Set(nameWords);
      const purpose = new Set(words(`${entry.useWhen} ${spec.hint}`));
      const detail = new Set(words(`${entry.syntax} ${entry.examples.join(' ')}`));
      const matched = terms.reduce(
        (total, term) =>
          total + (name.has(term) ? 20 : purpose.has(term) ? 4 : detail.has(term) ? 1 : 0),
        0,
      );
      // The task names the WHOLE command (e.g. "slide" in "add a slide"): a stronger signal than
      // sharing some words with a longer name (`add-table-slide` for the same task), which would
      // otherwise outscore the command the user actually asked for.
      const wholeName =
        nameWords.length > 0 &&
        nameWords.some((word) => termSet.has(word)) &&
        nameWords.every((word) => allTerms.has(word));
      const score = matched > 0 && wholeName ? matched + WHOLE_NAME_BONUS : matched;
      return { spec, entry, order, score };
    })
    .filter(({ score }) => score > 0)
    .sort((a, b) => b.score - a.score || a.order - b.order)
    .slice(0, COMMAND_DISCOVERY_LIMIT)
    .map(({ spec, entry }) => ({
      command: spec.usage.startsWith('/') ? `/${spec.verb}` : spec.verb,
      // Surface-specific signatures win over the generic help syntax (e.g. Word's bare read).
      syntax: spec.usage,
      useWhen: clip(entry.useWhen, 180),
      prerequisites: entry.discovery
        .filter((line) => allowed.has(line.split(/[\s/]/).find(Boolean) ?? ''))
        .slice(0, 2)
        .map((line) => clip(line, 120)),
      limits: [...entry.failureModes, ...entry.safety].slice(0, 3).map((line) => clip(line, 180)),
      example: exampleFor(spec, entry, terms),
    }));
}

/** A grammar spec that changes the document: a core write verb or a `/<kind>` specialized command. */
function isWriteSpec(spec: VerbSpec): boolean {
  return Object.hasOwn(WRITE_VERB_TO_KIND, spec.verb) || spec.usage.startsWith('/');
}

/**
 * One exact line for every write the surface advertises, in grammar order. A specialized command's
 * usage is only `/<kind> [key=value ...]`, which names no parameter, so its registry example is
 * shown instead. Every write is listed, not a task-ranked subset: a write the model has only seen by
 * name gets an invented syntax (live 2026-09-30: four different `chart` forms in four identical runs).
 */
export function writeSignatures(manifest: CapabilityManifest): string[] {
  return grammarFor(manifest)
    .filter(isWriteSpec)
    .map((spec) => (spec.usage.startsWith('/') ? signatureExample(manifest, spec) : spec.usage));
}

/**
 * A specialized command's first registry example that is one complete line of that same command.
 * Unlike {@link exampleFor}, a `<placeholder>` is allowed: this line is syntax, like a usage line.
 */
function signatureExample(manifest: CapabilityManifest, spec: VerbSpec): string {
  const example = helpFor(manifest, spec).examples.find(
    (value) =>
      !value.includes('\n') &&
      value.length <= 240 &&
      value.replace(/^\//, '').split(/\s/, 1)[0] === spec.verb,
  );
  return example ?? spec.usage;
}

export function renderCommandCard(card: CommandCard): string {
  return [
    `Command: ${card.command}`,
    `Syntax: ${card.syntax}`,
    `Use when: ${card.useWhen}`,
    ...(card.prerequisites.length ? [`If unresolved: ${card.prerequisites.join('; ')}`] : []),
    ...(card.limits.length ? [`Limits: ${card.limits.join(' ')}`] : []),
    `Example: ${card.example}`,
  ].join('\n');
}

function helpFor(manifest: CapabilityManifest, spec: VerbSpec): CommandHelpEntry {
  const entry = (COMMAND_HELP as Record<string, CommandHelpEntry>)[spec.verb];
  if (entry) return entry;
  const registry = registryEntryForKindAndSurface(spec.verb as ActuationKind, manifest.surface);
  return {
    command: spec.verb,
    syntax: spec.usage,
    useWhen: registry?.useWhen ?? spec.hint,
    discovery: registry?.discovery ?? [],
    sequence: [],
    examples: registry?.examples ?? [],
    doNot: [],
    failureModes: registry?.failureModes ?? [],
    safety: ['Live capabilities and host approval still apply.'],
  };
}

function exampleFor(spec: VerbSpec, entry: CommandHelpEntry, terms: string[]): string {
  // Do not turn another operation's example into an implied capability. Keep one complete line;
  // truncating a command would teach invalid syntax. A signature is safer than a cut payload.
  const example = entry.examples
    .filter(
      (value) =>
        !value.includes('\n') &&
        value.length <= 240 &&
        value.replace(/^\//, '').split(/\s/, 1)[0] === spec.verb &&
        !value.includes('<'),
    )
    .map((value, order) => {
      const vocabulary = new Set(words(value));
      return { value, order, score: terms.filter((term) => vocabulary.has(term)).length };
    })
    .sort((a, b) => b.score - a.score || a.order - b.order)[0]?.value;
  return example ?? spec.usage;
}

const STOP_WORDS = new Set(
  'a an and are as at be by can do for from have how i in is it me my need of on or please that the this to use want when with you'.split(
    ' ',
  ),
);

/** Colour names count as "color", so "make it blue" finds the command that changes colours. */
const COLOR_WORDS = new Set(
  'colour color red blue green yellow orange purple pink black white grey gray navy teal'.split(
    ' ',
  ),
);

function words(value: string): string[] {
  return [...new Set(value.toLowerCase().match(/[a-z0-9]+/g) ?? [])]
    .filter((word) => !STOP_WORDS.has(word))
    .map((word) => (COLOR_WORDS.has(word) ? 'color' : word))
    .map((word) => (word.length > 3 && word.endsWith('s') ? word.slice(0, -1) : word));
}

function clip(value: string, limit: number): string {
  return value.length <= limit ? value : `${value.slice(0, limit - 1)}…`;
}
