import {
  COMMAND_HELP,
  grammarFor,
  registryEntryForKindAndSurface,
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

/**
 * Deterministic lexical discovery over the same grammar/help used by execution. No model, network,
 * document content or dynamic registrations are involved. An unrelated query returns no cards;
 * the query itself is never reflected into the instruction channel.
 */
export function discoverCommands(manifest: CapabilityManifest, query: string): CommandCard[] {
  const specs = grammarFor(manifest);
  const terms = words(query.slice(0, 1024)).slice(0, 32);
  if (terms.length === 0) return [];
  const allowed = new Set(specs.map((spec) => spec.verb));
  return specs
    .map((spec, order) => {
      const entry = helpFor(manifest, spec);
      const name = new Set(words(spec.verb));
      const purpose = new Set(words(`${entry.useWhen} ${spec.hint}`));
      const detail = new Set(words(`${entry.syntax} ${entry.examples.join(' ')}`));
      const score = terms.reduce(
        (total, term) =>
          total + (name.has(term) ? 20 : purpose.has(term) ? 4 : detail.has(term) ? 1 : 0),
        0,
      );
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

function words(value: string): string[] {
  return [...new Set(value.toLowerCase().match(/[a-z0-9]+/g) ?? [])]
    .filter((word) => !STOP_WORDS.has(word))
    .map((word) => (word.length > 3 && word.endsWith('s') ? word.slice(0, -1) : word));
}

function clip(value: string, limit: number): string {
  return value.length <= limit ? value : `${value.slice(0, limit - 1)}…`;
}
