import {
  extractCommandBlock,
  isProgramCommand,
  isProgramExpr,
  parseProgramBlock,
  type ProgramEntry,
} from '@ge/contracts';

interface ParsedLine {
  line: string;
  actionable: boolean;
  explicitSingle: boolean;
}

/**
 * Extract an explicit user-authored CLI program from composer text.
 *
 * This is intentionally conservative:
 * - a fenced ```cmd block is always explicit and is returned verbatim;
 * - naked text must contain a contiguous run of real CLI lines;
 * - prose before/after that run is ignored, so "Populate this please" + many `set …` lines works;
 * - slash-style app intents (`/summarize`) are not treated as naked CLI specialized invokes.
 */
export function extractDirectCommandProgram(raw: string): string | undefined {
  const fenced = extractCommandBlock(raw);
  if (fenced !== null) {
    const body = normalizeProgram(fenced);
    return body.length > 0 ? body : undefined;
  }

  const lines = raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  if (lines.length === 0) return undefined;

  let best: ParsedLine[] = [];
  let current: ParsedLine[] = [];
  for (const line of lines) {
    const parsed = parseNakedCliLine(line);
    // A set-shaped sentence must reach the planner in full, not vanish from a partial CLI run.
    if (parsed === 'ambiguous-set') return undefined;
    if (parsed) {
      current.push(parsed);
      continue;
    }
    if (current.length > best.length) best = current;
    current = [];
  }
  if (current.length > best.length) best = current;

  const actionableCount = best.filter((line) => line.actionable).length;
  if (best.length >= 2 && actionableCount > 0) return best.map((line) => line.line).join('\n');
  if (best.length === 1 && lines.length === 1 && best[0]?.actionable && best[0].explicitSingle) {
    return best[0].line;
  }
  return undefined;
}

function parseNakedCliLine(line: string): ParsedLine | 'ambiguous-set' | undefined {
  const { found, entries } = parseProgramBlock(`\`\`\`cmd\n${line}\n\`\`\``);
  if (!found || entries.length !== 1) return undefined;
  const entry = entries[0]!;
  if ('error' in entry) return undefined;
  if (!isNakedDirectEntry(entry)) return undefined;
  if (isProgramCommand(entry) && entry.verb === 'set' && readsAsProse(line, entry))
    return 'ambiguous-set';
  // Read/search/outline overlap ordinary English. Bare read syntax must be recognisably
  // addressed; arbitrary selectors and search expressions remain available in explicit fences.
  if (isProgramCommand(entry)) {
    if (
      entry.verb === 'read' &&
      !/^(?:read|read\s+(?:selection|document(?::[\d:]+)?|(?:'[^']+'|[\w.-]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|slide:\d+))$/i.test(
        line,
      )
    )
      return undefined;
    if (entry.verb === 'search') return undefined;
    if (entry.verb === 'outline' && line.toLowerCase() !== 'outline') return undefined;
  }
  return {
    line,
    actionable: isActionableEntry(entry),
    explicitSingle: isExplicitSingleLine(entry) && !readsAsProse(line, entry),
  };
}

/**
 * A lone line whose arguments are mostly plain words is a sentence that happens to start with a
 * CLI verb ("Read the last row for me", "Search for the payment clause"), not a command. CLI
 * arguments are refs (`A1:B9`, `slide:2`, `result:x`), key=value props or quoted strings, so
 * count only the bare alphabetic words left after the verb once quoted spans are removed.
 * Naked `set` text values must be quoted; otherwise `set A1 to ...` writes prose as a value.
 */
function readsAsProse(line: string, entry: ProgramEntry): boolean {
  const set = isProgramCommand(entry) && entry.verb === 'set';
  if (
    set &&
    (entry.valueExpr || entry.value.startsWith('=') || /^(true|false|null)$/i.test(entry.value))
  )
    return false;
  const args = line
    .replace(/"(?:[^"\\]|\\.)*"|'[^']*'/g, ' ')
    .trim()
    .split(/\s+/)
    .slice(1);
  return args.filter((token) => /^[a-z]+[.,!?]?$/i.test(token)).length >= (set ? 1 : 2);
}

function isNakedDirectEntry(entry: ProgramEntry): boolean {
  if (isProgramExpr(entry)) return true;
  if (!isProgramCommand(entry)) return false;
  // `/summarize @this` and friends are composer intents, not direct CLI specialized invocations.
  if (entry.verb === 'invoke') return false;
  // `find` takes a workspace path. "Find and replace X with Y" also parses (path "and"), so a typed
  // `find` runs directly only when its first argument looks like a path; otherwise it is prose.
  if (entry.verb === 'find') return /^[/.~]/.test(entry.path);
  return true;
}

function isActionableEntry(entry: ProgramEntry): boolean {
  if (isProgramExpr(entry)) return true;
  if (!isProgramCommand(entry)) return false;
  return entry.verb !== 'done' && entry.verb !== 'help';
}

function isExplicitSingleLine(entry: ProgramEntry): boolean {
  if (isProgramExpr(entry)) return true;
  if (!isProgramCommand(entry)) return false;
  if (entry.verb === 'set') return /(?:^|!)\$?[A-Z]{1,3}\$?\d+\b/i.test(entry.cell);
  return entry.verb !== 'done' && entry.verb !== 'help';
}

function normalizeProgram(program: string): string {
  return program
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
    .join('\n');
}
