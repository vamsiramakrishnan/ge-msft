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

function parseNakedCliLine(line: string): ParsedLine | undefined {
  const { found, entries } = parseProgramBlock(`\`\`\`cmd\n${line}\n\`\`\``);
  if (!found || entries.length !== 1) return undefined;
  const entry = entries[0]!;
  if ('error' in entry) return undefined;
  if (!isNakedDirectEntry(entry)) return undefined;
  return {
    line,
    actionable: isActionableEntry(entry),
    explicitSingle: isExplicitSingleLine(entry),
  };
}

function isNakedDirectEntry(entry: ProgramEntry): boolean {
  if (isProgramExpr(entry)) return true;
  if (!isProgramCommand(entry)) return false;
  // `/summarize @this` and friends are composer intents, not direct CLI specialized invocations.
  return entry.verb !== 'invoke';
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
