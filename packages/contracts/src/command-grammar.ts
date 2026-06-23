import { z } from 'zod';
import type { ActuationKind, CapabilityManifest } from './capability.js';
import {
  isExpressionLine,
  parseExpressionLine,
  type ExprParseError,
  type ParsedExpr,
} from './expr-grammar.js';

/**
 * ADR-0004 — the command-line protocol grammar (the single source of truth).
 *
 * The grounded `streamAssist` model emits flat command lines inside a fenced ` ```cmd `
 * block (it prepends a `**thought**` preamble — we ignore the prose and parse only the
 * fence). The runtime parses → validates → compiles each line into the *existing* typed
 * boundary objects (`ActuationRequest` / Layer-B read calls). This module owns the grammar
 * and the parser; the runtime owns the compile step (see `@ge/runtime` command-protocol).
 *
 * The command line is the *assembly language* the model emits; the typed `ActuationRequest`
 * (ADR-0002) is the *bytecode* the bridge executes. Robust quoting is now OUR job — the
 * model never has to emit schema-valid JSON. On a parse failure we return a CLI-style
 * corrective error (`unknown verb "writ" — did you mean "write"? (run help)`) the model
 * self-corrects on the next turn.
 *
 * SCOPE: `outline · read · search · set · suggest · comment · format · done · help`.
 */

/** Read verbs (Layer-B host reads, ADR-0003). Always advertised; never gated. */
export const READ_VERBS = ['outline', 'read', 'search'] as const;

/** Control verbs. Always advertised; not actuations. */
export const CONTROL_VERBS = ['done', 'help'] as const;

/**
 * Write verbs → the `ActuationKind` (ADR-0002) they compile to. A write verb is advertised
 * for a surface ONLY when the surface's `CapabilityManifest.actuations[]` includes its mapped
 * kind (so `set` shows for Excel's `write-cells`, `suggest` for Word's `tracked-change`). The
 * parser, the advertisement, and the runtime compiler all derive from this single map.
 */
export const WRITE_VERB_TO_KIND = {
  set: 'write-cells',
  suggest: 'tracked-change',
  comment: 'add-comment',
  format: 'format-cells',
  reply: 'comment-reply',
} satisfies Record<string, ActuationKind>;

export type WriteVerb = keyof typeof WRITE_VERB_TO_KIND;
export type ReadVerb = (typeof READ_VERBS)[number];
export type ControlVerb = (typeof CONTROL_VERBS)[number];

/** Every verb this slice understands (for did-you-mean + advertisement). */
export const ALL_VERBS = [
  ...READ_VERBS,
  ...CONTROL_VERBS,
  ...(Object.keys(WRITE_VERB_TO_KIND) as WriteVerb[]),
] as const;

/* ───────────────────────────── ParsedCommand ───────────────────────────── */

/**
 * A successfully parsed command line — a discriminated union on `verb`. The selector/args
 * are surface-agnostic strings here; the runtime compiler maps them onto surface-specific
 * targets (`set` cell → `target.range`; `suggest` oldText → `target.matchText`).
 */
export type ParsedCommand =
  | { verb: 'outline' }
  | { verb: 'read'; selector: string }
  | { verb: 'search'; text: string }
  | { verb: 'set'; cell: string; value: string }
  | { verb: 'suggest'; oldText: string; newText: string }
  | { verb: 'comment'; selector: string; text: string }
  | { verb: 'format'; range: string; props: Record<string, string> }
  | { verb: 'reply'; commentId: string; text: string }
  | { verb: 'done' }
  | { verb: 'help' };

export const ParsedCommandSchema: z.ZodType<ParsedCommand> = z.discriminatedUnion('verb', [
  z.object({ verb: z.literal('outline') }),
  z.object({ verb: z.literal('read'), selector: z.string() }),
  z.object({ verb: z.literal('search'), text: z.string() }),
  z.object({ verb: z.literal('set'), cell: z.string(), value: z.string() }),
  z.object({ verb: z.literal('suggest'), oldText: z.string(), newText: z.string() }),
  z.object({ verb: z.literal('comment'), selector: z.string(), text: z.string() }),
  z.object({ verb: z.literal('format'), range: z.string(), props: z.record(z.string()) }),
  z.object({ verb: z.literal('reply'), commentId: z.string(), text: z.string() }),
  z.object({ verb: z.literal('done') }),
  z.object({ verb: z.literal('help') }),
]);

/** A parse error carries a CLI-style corrective message the model self-corrects against. */
export interface CommandParseError {
  error: string;
}

export function isCommandParseError(c: ParsedCommand | CommandParseError): c is CommandParseError {
  return 'error' in c;
}

/* ───────────────────────────── fence extraction ───────────────────────── */

/**
 * Pull the contents of the ```cmd fenced block out of the model's reply, ignoring the
 * surrounding `**thought**` prose. Returns the inner text (trimmed) or `null` when there is
 * no fenced cmd block (the runtime treats that as a re-prompt, not an error — ADR-0004 §3.2).
 *
 * Tolerant of a language tag with trailing spaces and of CRLF line endings.
 */
export function extractCommandBlock(modelText: string): string | null {
  // Non-greedy capture of the first ```cmd … ``` fence. `[^\S\n]` = horizontal whitespace,
  // so a `cmd` tag with trailing spaces still matches but a different language tag does not.
  const match = modelText.match(/```cmd[^\S\n]*\r?\n([\s\S]*?)```/i);
  return match ? match[1]!.trim() : null;
}

/* ───────────────────────────── line parsing ───────────────────────────── */

/**
 * Parse ONE command line into a `ParsedCommand` or a corrective `{ error }`.
 *
 * The first whitespace-delimited token is the verb (case-insensitive); arg parsing is
 * verb-specific. This is the whole reliability thesis — robust quoting is our job:
 *   • `set <cell> <value>` — `cell` is the first token; `value` is the FULL remainder
 *     (may contain spaces and commas, e.g. `=SUM(A1, A2)`) — never re-split.
 *   • `suggest "old" => "new"` — two double-quoted strings (with `\"`/`\\` escapes),
 *     separated by `=>` (surrounding spaces tolerated; `->` also accepted).
 *   • `comment <selector> "text"` — surface-portable: a bare selector (Excel cell) OR a
 *     quoted anchor (Word content anchor) for the first arg, then a quoted comment body.
 *   • `format <range> k=v k=v ...` — first token is the range; the rest are `key=value`
 *     pairs (split on the FIRST `=`; values may contain `# $ , . %`, no quotes needed).
 *   • `read`/`search` — the remainder is the selector / search text (verbatim).
 *   • `outline`/`done`/`help` — no args.
 * An unknown/garbled verb yields a did-you-mean against the advertised verbs.
 */
export function parseCommandLine(line: string): ParsedCommand | CommandParseError {
  const trimmed = line.trim();
  // Caller skips blanks/comments, but be defensive: treat them as nothing-to-parse.
  if (trimmed === '' || trimmed.startsWith('#')) {
    return { error: 'empty command line' };
  }

  const firstSpace = trimmed.search(/\s/);
  const rawVerb = firstSpace === -1 ? trimmed : trimmed.slice(0, firstSpace);
  const verb = rawVerb.toLowerCase();
  const rest = firstSpace === -1 ? '' : trimmed.slice(firstSpace + 1).trim();

  switch (verb) {
    case 'outline':
      return { verb: 'outline' };
    case 'done':
      return { verb: 'done' };
    case 'help':
      return { verb: 'help' };

    case 'read': {
      // Excel: an A1/NamedRange selector. Word: whole-doc (no selector) — empty string is fine,
      // the runtime read-intent treats an empty selector as "whole document".
      return { verb: 'read', selector: rest };
    }

    case 'search': {
      if (rest === '') return { error: 'search needs text — usage: search <text>' };
      // Tolerate the model wrapping the query in quotes.
      return { verb: 'search', text: stripWrappingQuotes(rest) };
    }

    case 'set': {
      const sp = rest.search(/\s/);
      if (sp === -1) {
        return { error: 'set needs a cell and a value — usage: set <A1> <value|=formula>' };
      }
      const cell = rest.slice(0, sp);
      const value = rest.slice(sp + 1).trim();
      if (cell === '' || value === '') {
        return { error: 'set needs a cell and a value — usage: set <A1> <value|=formula>' };
      }
      // value is the FULL remainder — commas/spaces/`=SUM(A1, A2)` preserved verbatim.
      return { verb: 'set', cell, value };
    }

    case 'suggest': {
      const parsed = parseSuggest(rest);
      return parsed;
    }

    case 'comment': {
      const parsed = parseComment(rest);
      return parsed;
    }

    case 'format': {
      const parsed = parseFormat(rest);
      return parsed;
    }

    case 'reply': {
      const parsed = parseReply(rest);
      return parsed;
    }

    default:
      return { error: unknownVerbError(verb) };
  }
}

/**
 * `suggest "old text" => "new text"` — two double-quoted strings separated by `=>` (or `->`),
 * each supporting `\"` and `\\` escapes. We scan the first quoted string, require a `=>`/`->`
 * separator, then scan the second. Surrounding spaces are tolerated throughout.
 */
function parseSuggest(rest: string): ParsedCommand | CommandParseError {
  const usage = 'suggest needs two quoted strings — usage: suggest "old text" => "new text"';
  const first = scanQuoted(rest, 0);
  if (!first) return { error: usage };

  // Between the two strings: optional spaces, then `=>` or `->`.
  let i = first.end;
  while (i < rest.length && /\s/.test(rest[i]!)) i++;
  if (rest.startsWith('=>', i)) i += 2;
  else if (rest.startsWith('->', i)) i += 2;
  else return { error: usage };
  while (i < rest.length && /\s/.test(rest[i]!)) i++;

  const second = scanQuoted(rest, i);
  if (!second) return { error: usage };

  // Anything trailing the closing quote (other than whitespace) is malformed.
  if (rest.slice(second.end).trim() !== '') return { error: usage };

  if (first.value === '') return { error: 'suggest old text cannot be empty' };
  return { verb: 'suggest', oldText: first.value, newText: second.value };
}

/**
 * `comment <selector> "text"` — surface-portable across Excel (cell selector) and Word (content
 * anchor). The first argument is EITHER a bare, unquoted selector (e.g. `Sales!A16`) OR a quoted
 * anchor string (Word, supporting `\"`/`\\` escapes); the second argument is the quoted comment
 * body. The selector goes in `selector`, the body in `text`.
 */
function parseComment(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'comment needs a selector/anchor and a quoted comment — usage: comment <cell> "text"  OR  comment "anchor" "text"';

  // Skip leading whitespace.
  let i = 0;
  while (i < rest.length && /\s/.test(rest[i]!)) i++;
  if (i >= rest.length) return { error: usage };

  // First arg: a quoted anchor (Word) or a bare selector token (Excel).
  let selector: string;
  if (rest[i] === '"') {
    const anchor = scanQuoted(rest, i);
    if (!anchor) return { error: usage };
    selector = anchor.value;
    i = anchor.end;
  } else {
    const sp = rest.slice(i).search(/\s/);
    if (sp === -1) return { error: usage }; // selector but no comment body
    selector = rest.slice(i, i + sp);
    i += sp;
  }

  // Separator whitespace, then the quoted comment body.
  while (i < rest.length && /\s/.test(rest[i]!)) i++;
  const body = scanQuoted(rest, i);
  if (!body) return { error: usage };

  // Anything after the closing quote (other than whitespace) is malformed.
  if (rest.slice(body.end).trim() !== '') return { error: usage };

  if (selector === '') return { error: usage };
  return { verb: 'comment', selector, text: body.value };
}

/**
 * `format <range> k=v k=v ...` — first token is the A1/NamedRange; the rest are `key=value`
 * pairs, each split on the FIRST `=` only so values may carry `# $ , . %` unquoted
 * (e.g. `fill=#FFF2CC numberFormat=$#,##0.00 bold=true`). A format with no range or no props is
 * a corrective error.
 */
function parseFormat(rest: string): ParsedCommand | CommandParseError {
  const usage = 'format needs a range and at least one key=value — usage: format <range> k=v ...';
  const tokens = rest.split(/\s+/).filter((t) => t !== '');
  if (tokens.length === 0) return { error: usage };

  const range = tokens[0]!;
  const pairs = tokens.slice(1);
  if (pairs.length === 0) return { error: usage };

  const props: Record<string, string> = {};
  for (const pair of pairs) {
    const eq = pair.indexOf('=');
    if (eq <= 0) {
      return {
        error: `format expects key=value pairs — got "${pair}" (usage: format <range> k=v)`,
      };
    }
    const key = pair.slice(0, eq);
    const value = pair.slice(eq + 1);
    props[key] = value;
  }

  return { verb: 'format', range, props };
}

/**
 * `reply <commentId> "text"` (ADR-0006 `comment-reply`). The first bare token is the comment id
 * (host-opaque, e.g. `{xyz}` / a GUID — no spaces); the second argument is the quoted reply body
 * (with `\"`/`\\` escapes via {@link scanQuoted}). Gated behind the `comment-reply` actuation,
 * which Word/Excel advertise. A missing id or body is a corrective error.
 */
function parseReply(rest: string): ParsedCommand | CommandParseError {
  const usage = 'reply needs a comment id and a quoted reply — usage: reply <commentId> "text"';

  // Skip leading whitespace.
  let i = 0;
  while (i < rest.length && /\s/.test(rest[i]!)) i++;
  if (i >= rest.length) return { error: usage };

  // First arg: a bare comment-id token (no quoting — host ids carry no spaces).
  const sp = rest.slice(i).search(/\s/);
  if (sp === -1) return { error: usage }; // id but no reply body
  const commentId = rest.slice(i, i + sp);
  i += sp;

  // Separator whitespace, then the quoted reply body.
  while (i < rest.length && /\s/.test(rest[i]!)) i++;
  const body = scanQuoted(rest, i);
  if (!body) return { error: usage };

  // Anything after the closing quote (other than whitespace) is malformed.
  if (rest.slice(body.end).trim() !== '') return { error: usage };

  if (commentId === '') return { error: usage };
  return { verb: 'reply', commentId, text: body.value };
}

/**
 * Scan a double-quoted string starting at `start` (which must index the opening `"`).
 * Honors `\"` (literal quote) and `\\` (literal backslash). Returns the unescaped value and
 * the index just past the closing quote, or `null` if there is no well-formed quoted string.
 */
export function scanQuoted(s: string, start: number): { value: string; end: number } | null {
  if (s[start] !== '"') return null;
  let value = '';
  let i = start + 1;
  while (i < s.length) {
    const ch = s[i]!;
    if (ch === '\\' && i + 1 < s.length) {
      const next = s[i + 1]!;
      value += next === '"' || next === '\\' ? next : `\\${next}`;
      i += 2;
      continue;
    }
    if (ch === '"') return { value, end: i + 1 };
    value += ch;
    i++;
  }
  return null; // unterminated
}

/** Strip a single pair of wrapping single/double quotes, if present. */
function stripWrappingQuotes(s: string): string {
  if (s.length >= 2) {
    const a = s[0]!;
    const b = s[s.length - 1]!;
    if ((a === '"' && b === '"') || (a === "'" && b === "'")) return s.slice(1, -1);
  }
  return s;
}

/* ───────────────────────────── block parsing ──────────────────────────── */

/**
 * Parse the whole model reply: extract the ```cmd fence, then parse each non-blank, non-comment
 * line. `found` is false when there is no fence (→ the runtime re-prompts). Comments (`#…`) and
 * blank lines are skipped, never errors.
 */
export function parseCommandBlock(modelText: string): {
  found: boolean;
  commands: Array<ParsedCommand | CommandParseError>;
} {
  const block = extractCommandBlock(modelText);
  if (block === null) return { found: false, commands: [] };

  const commands: Array<ParsedCommand | CommandParseError> = [];
  for (const raw of block.split('\n')) {
    const line = raw.trim();
    if (line === '' || line.startsWith('#')) continue;
    commands.push(parseCommandLine(line));
  }
  return { found: true, commands };
}

/* ───────────────────────────── program parsing (ADR-0005) ───────────────── */

/**
 * One entry in a parsed program block: a simple ADR-0004 command, an ADR-0005 expression
 * (pipeline / `let`), or a corrective parse error from either layer. The runtime loop dispatches
 * on `kind` (expression) vs `verb` (command) vs `error`.
 */
export type ProgramEntry = ParsedCommand | ParsedExpr | CommandParseError;

export function isProgramExpr(entry: ProgramEntry): entry is ParsedExpr {
  return 'kind' in entry && (entry.kind === 'pipeline' || entry.kind === 'let');
}

export function isProgramCommand(entry: ProgramEntry): entry is ParsedCommand {
  return 'verb' in entry;
}

/**
 * Parse the whole model reply as an ADR-0005 *program*: extract the ```cmd fence, then for each
 * non-blank, non-comment line route to the expression parser (when the line `isExpressionLine` —
 * a top-level `|` or a leading `let`) or to the unchanged ADR-0004 command parser otherwise.
 *
 * This is a superset of {@link parseCommandBlock}: every line that is NOT an expression parses
 * EXACTLY as before, so all ADR-0004 behavior is preserved. `found` is false when there is no
 * fence (the runtime re-prompts).
 */
export function parseProgramBlock(modelText: string): {
  found: boolean;
  entries: ProgramEntry[];
} {
  const block = extractCommandBlock(modelText);
  if (block === null) return { found: false, entries: [] };

  const entries: ProgramEntry[] = [];
  for (const raw of block.split('\n')) {
    const line = raw.trim();
    if (line === '' || line.startsWith('#')) continue;
    entries.push(isExpressionLine(line) ? parseExpressionLine(line) : parseCommandLine(line));
  }
  return { found: true, entries };
}

export type { ExprParseError };

/* ───────────────────────── capability-scoped advertisement ─────────────── */

/** One advertised verb: its name, a usage line, and a one-line hint for the prompt. */
export interface VerbSpec {
  verb: string;
  usage: string;
  hint: string;
}

/**
 * The capability-scoped grammar advertisement for a surface. Control verbs are always advertised; a
 * READ verb (`outline`/`read`/`search`) appears ONLY when it is in `manifest.reads` (ADR-0006 — a
 * surface must never advertise a read it cannot serve), and a WRITE verb appears ONLY when
 * `manifest.actuations[]` contains its mapped `ActuationKind`. Surface selector hints differ (Excel
 * reads an A1/NamedRange; Word's `read` is whole-document), so the smaller per-surface grammar is
 * fewer tokens to get wrong.
 */
export function grammarFor(manifest: CapabilityManifest): VerbSpec[] {
  const isExcelLike = manifest.surface === 'excel';
  const readSelector = isExcelLike
    ? { verb: 'read', usage: 'read <A1|NamedRange>', hint: 'read a range, e.g. read Sales!C2:C7' }
    : { verb: 'read', usage: 'read', hint: 'read the whole document' };

  // Read verbs, scoped by manifest.reads (ADR-0006). Advertise a read verb only when the surface
  // declares it serves that read — otherwise the grammar would advertise an unreachable read.
  const declaredReads = new Set(manifest.reads ?? []);
  const readSpecByVerb: Record<ReadVerb, VerbSpec> = {
    outline: { verb: 'outline', usage: 'outline', hint: 'show the document/workbook structure' },
    read: readSelector,
    search: { verb: 'search', usage: 'search <text>', hint: 'find content containing the text' },
  };

  const specs: VerbSpec[] = [];
  for (const verb of READ_VERBS) {
    if (declaredReads.has(verb)) specs.push(readSpecByVerb[verb]);
  }

  // Write verbs, gated by the advertised actuation kinds. Derived from WRITE_VERB_TO_KIND so a
  // new (deferred) write verb only needs an entry there + its kind in the manifest.
  const kinds = new Set(manifest.actuations.map((a) => a.kind));
  for (const [verb, kind] of Object.entries(WRITE_VERB_TO_KIND) as [WriteVerb, ActuationKind][]) {
    if (!kinds.has(kind)) continue;
    specs.push(writeVerbSpec(verb, isExcelLike));
  }

  specs.push(
    { verb: 'done', usage: 'done', hint: 'finish — you have completed the task' },
    { verb: 'help', usage: 'help', hint: 'list the available commands' },
  );
  return specs;
}

/**
 * The usage/hint for a write verb, surface-aware where the selector differs. `comment` reads a
 * bare cell on Excel and a quoted content anchor on Word; `format` is Excel-only (gated by the
 * `format-cells` actuation, which only Excel advertises). Kept beside WRITE_VERB_TO_KIND.
 */
function writeVerbSpec(verb: WriteVerb, isExcelLike: boolean): VerbSpec {
  switch (verb) {
    case 'set':
      return {
        verb: 'set',
        usage: 'set <A1> <value|=formula>',
        hint: 'write one cell, e.g. set Sales!F2 =C2-D2',
      };
    case 'suggest':
      return {
        verb: 'suggest',
        usage: 'suggest "old text" => "new text"',
        hint: 'propose a tracked change anchored on the exact existing text',
      };
    case 'comment':
      return isExcelLike
        ? {
            verb: 'comment',
            usage: 'comment <cell> "text"',
            hint: 'comment on a cell, e.g. comment Sales!A16 "anomalous spike"',
          }
        : {
            verb: 'comment',
            usage: 'comment "anchor" "text"',
            hint: 'comment anchored on the exact existing text',
          };
    case 'format':
      return {
        verb: 'format',
        usage: 'format <range> k=v ...',
        hint: 'format a range, e.g. format Sales!A16:C16 bold=true fill=#FFF2CC numberFormat=$#,##0.00',
      };
    case 'reply':
      return {
        verb: 'reply',
        usage: 'reply <commentId> "text"',
        hint: 'reply to an existing comment by its id, e.g. reply {3f2a} "addressed in the redline"',
      };
  }
}

/* ───────────────────────────── did-you-mean ───────────────────────────── */

/**
 * A CLI-style corrective for an unknown verb, with a Levenshtein-nearest suggestion against the
 * advertised verbs when one is close enough.
 */
function unknownVerbError(verb: string): string {
  const suggestion = nearestVerb(verb);
  const tail = suggestion ? ` — did you mean "${suggestion}"? (run help)` : ' (run help)';
  return `unknown verb "${verb}"${tail}`;
}

/** The closest advertised verb within an edit-distance threshold, or undefined. */
function nearestVerb(verb: string): string | undefined {
  let best: string | undefined;
  let bestDist = Infinity;
  for (const candidate of ALL_VERBS) {
    const d = levenshtein(verb, candidate);
    if (d < bestDist) {
      bestDist = d;
      best = candidate;
    }
  }
  // Only suggest when the typo is plausibly the same word: within edits ≤ ~⌈len/2⌉, capped at 3.
  const threshold = Math.min(3, Math.max(1, Math.ceil(verb.length / 2)));
  return best !== undefined && bestDist <= threshold ? best : undefined;
}

/** Classic iterative Levenshtein edit distance. */
function levenshtein(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  let curr = new Array<number>(n + 1);
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j]! + 1, curr[j - 1]! + 1, prev[j - 1]! + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[n]!;
}
