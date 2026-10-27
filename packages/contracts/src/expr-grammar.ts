import { z } from 'zod';
import { scanQuoted } from './command-grammar.js';

/**
 * The ADR-0004 effect verbs — they cannot be a pipeline source in Phase 1 (pure-only). Inlined
 * (not imported from `WRITE_VERB_TO_KIND`) to avoid a module-eval-time dependency on the
 * command-grammar module (the two modules reference each other), and kept in sync by the
 * exhaustiveness test in `expr-grammar.test.ts`.
 */
export const EFFECT_VERBS: ReadonlySet<string> = new Set([
  'set',
  'grid',
  'suggest',
  'comment',
  'format',
  'reply',
  // ADR-0006 CLI parity effect verbs (insert-slide / append-page / reply-mail / post-message).
  'slide',
  'page',
  'mail',
  'post',
  // Outlook create-mail (compose a new draft) — also an effect, never a pipeline source.
  'compose',
  // ADR-0007 host-native Excel kinds — effects, never a pipeline source.
  'table',
  'chart',
  'cf',
  'shape',
  // `spill` consumes a table expression but is itself an effect terminal (you can't pipe out of it).
  'spill',
]);

/** The Phase-1 corrective when the model tries to compose an effect. */
export const EFFECT_COMPOSE_ERROR =
  "effects can't be composed yet — emit them as standalone commands (Phase 2)";

/**
 * The pure pipeline transforms, by name — the canonical list the composition grammar advertises.
 * The runtime's `TRANSFORMS` registry (packages/runtime/src/compose.ts) MUST provide exactly these
 * keys; a drift test in `compose.test.ts` binds the two. Kept here (not in runtime) so the boundary
 * owns the name set and `shadowsBuiltin` can refuse a skill that would shadow a transform.
 */
export const TRANSFORM_NAMES = [
  'filter',
  'select',
  'sum',
  'avg',
  'min',
  'max',
  'count',
  'sort',
  'head',
  'tail',
  'sed',
  'derive',
  // C1 richer algebra — relational transforms over bound tables.
  'join',
  'group',
  'distinct',
  'rename',
] as const;

/**
 * ADR-0005 (Phase 1) — the expression layer over the ADR-0004 command grammar.
 *
 * The keystone is a typed *value* between reads and transforms: reads produce values; pure
 * transforms compose them. This module is the **structural** parser only — it turns an
 * expression line into a `ParsedExpr` AST. It deliberately does NOT know transform names or
 * argument shapes: a stage is `{ name, args }` with a RAW arg string; the runtime owns the
 * transform registry, validates names, and parses per-transform args. Contracts stays the
 * single source of truth for *shape*, the runtime for *meaning*.
 *
 * Phase 1 is PURE only — a pipeline composes reads through transforms into a value; `let` binds
 * a value to a `$name`. There is NO piping into effects (`set`/`suggest`/`comment`/`format`);
 * the runtime rejects that with a corrective error. Any line that is NOT an expression (no
 * top-level ` | ` and not starting with `let `) is left to the ADR-0004 simple-command parser,
 * unchanged.
 *
 * SCOPE: `<source> ( '|' <stage> )*` and `let $name = <pipeline>`, where
 *   <source> = `read <selector>` | `search <text>` | `outline` | `$var`
 *            | `<number>` | `"<text>"` | `table [[<header>…], [<row>…], …]`   (C1 literals).
 */

/* ───────────────────────────── AST ───────────────────────────── */

/** The head of a pipeline — where the value comes from before any transform runs. */
export type PipeSource =
  | { src: 'read'; selector: string }
  | { src: 'search'; text: string }
  | { src: 'outline' }
  | { src: 'var'; name: string }
  // C1 literal sources — a pipeline (or a `let`) may start from a value written in the program
  // itself: `let $n = 5`, `let $s = "text"`, `let $t = table [["a","b"],["1","2"]]`. The table
  // literal is a JSON array of rows whose first row is the header; cells are strings (numbers and
  // booleans are stringified). Literals are data the MODEL wrote, so they carry no host content.
  | { src: 'number'; value: number }
  | { src: 'text'; value: string }
  | { src: 'table'; columns: string[]; rows: string[][] };

/** A pipeline source that reaches the host (everything but a `$var` lookup or a literal). */
export type HostPipeSource = Extract<PipeSource, { src: 'read' | 'search' | 'outline' }>;

/**
 * One transform stage: the first token is the transform NAME; the remainder is the verbatim
 * `args` string (the runtime parses it per-transform, e.g. `filter region=East`, `sum amount`,
 * `sort total desc`). Empty `args` is allowed (e.g. `count`).
 */
export interface TransformStage {
  name: string;
  args: string;
}

/** A pipeline expression: a source folded through zero-or-more transform stages. */
export interface PipelineExpr {
  kind: 'pipeline';
  source: PipeSource;
  stages: TransformStage[];
}

/** A binding: evaluate a pipeline and bind the resulting value to `$name` in the env. */
export interface LetExpr {
  kind: 'let';
  name: string;
  pipeline: PipelineExpr;
}

/** A parsed expression line — either a bare pipeline or a `let` binding of one. */
export type ParsedExpr = PipelineExpr | LetExpr;

/* ───────────────────────────── Zod ───────────────────────────── */

export const PipeSourceSchema: z.ZodType<PipeSource> = z.discriminatedUnion('src', [
  z.object({ src: z.literal('read'), selector: z.string() }),
  z.object({ src: z.literal('search'), text: z.string() }),
  z.object({ src: z.literal('outline') }),
  z.object({ src: z.literal('var'), name: z.string() }),
  z.object({ src: z.literal('number'), value: z.number() }),
  z.object({ src: z.literal('text'), value: z.string() }),
  z.object({
    src: z.literal('table'),
    columns: z.array(z.string()),
    rows: z.array(z.array(z.string())),
  }),
]);

export const TransformStageSchema: z.ZodType<TransformStage> = z.object({
  name: z.string(),
  args: z.string(),
});

export const PipelineExprSchema: z.ZodType<PipelineExpr> = z.object({
  kind: z.literal('pipeline'),
  source: PipeSourceSchema,
  stages: z.array(TransformStageSchema),
});

export const LetExprSchema: z.ZodType<LetExpr> = z.object({
  kind: z.literal('let'),
  name: z.string(),
  pipeline: PipelineExprSchema,
});

export const ParsedExprSchema: z.ZodType<ParsedExpr> = z.union([PipelineExprSchema, LetExprSchema]);

export function isParsedExpr(node: { kind?: string }): node is ParsedExpr {
  return node.kind === 'pipeline' || node.kind === 'let';
}

/* ───────────────────────── line classification ─────────────────────────── */

/**
 * Is this line an expression (handled here) rather than an ADR-0004 simple command? True when it
 * starts with `let ` OR contains a top-level pipe ` | ` (a pipe not inside a double-quoted span,
 * so a quoted comment body with a literal `|` doesn't false-trigger). A plain `read X`,
 * `outline`, `search foo`, or `set …` line is NOT an expression and falls through to the simple
 * parser exactly as today.
 */
export function isExpressionLine(line: string): boolean {
  const trimmed = line.trim();
  if (trimmed === '' || trimmed.startsWith('#')) return false;
  if (/^let\s/i.test(trimmed)) return true;
  return findTopLevelPipes(trimmed).length > 0;
}

/* ───────────────────── effect-arg expressions (ADR-0005 Phase 2) ──────────── */

/**
 * ADR-0005 Phase 2 — an effect verb's value/text argument may be an EXPRESSION evaluated against
 * the binding env, instead of a literal. The rule is deliberately unambiguous (so back-compat is
 * total): an effect arg is an expression IFF it is *exactly*
 *   • a bare `$var` (e.g. `set B3 = $total`), or
 *   • a parenthesized pipeline `( <pipeline> )` (e.g. `set Summary!B2 = ($anz | sum Revenue)`).
 * Anything else — plain text, a `=formula`, a value with stray parens — is a LITERAL exactly as
 * ADR-0004. The expression's `$var` source is allowed (and is the keystone connection); the
 * runtime evaluates it via {@link evalExpr} at dry-run, NEVER actuating during evaluation.
 *
 * The parenthesized form is parsed through the SAME pure {@link parsePipeline}/{@link parseSource}
 * path as a top-level pipeline, so an effect arg can read+compute but can NEVER smuggle a write: a
 * pipe INTO or out of an effect verb (`$x | set …`, `read X | set …`) is rejected with
 * {@link EFFECT_COMPOSE_ERROR} by the existing pure-only guard. Returns `undefined` when the arg is
 * a literal (the caller keeps the verbatim literal), or a `ParsedExpr` / `ExprParseError`.
 */
export function parseEffectArg(arg: string): ParsedExpr | ExprParseError | undefined {
  // The model writes the expression form with an assignment `=`: `set B3 = $total`,
  // `set B2 = ($a | sum X)`. Strip a leading `=` ONLY when it is followed by whitespace (the
  // assignment form) so a literal `=formula` (`=SUM(A1,A2)`, `=C2-D2` — no space after `=`) is
  // never mistaken for an expression. Comment/reply bodies have no `=`; they pass through unchanged.
  let trimmed = arg.trim();
  if (/^=\s/.test(trimmed)) trimmed = trimmed.slice(1).trim();

  // Bare `$var` → a one-stage-less pipeline whose source is the binding.
  if (isVarName(trimmed)) {
    return { kind: 'pipeline', source: { src: 'var', name: stripDollar(trimmed) }, stages: [] };
  }
  // A fully-parenthesized pipeline `( … )`: the FIRST char is `(` and its matching close is the
  // LAST char (so a literal that merely contains parens — `=SUM(A1, A2)` — is NOT an expression).
  if (trimmed.startsWith('(')) {
    const close = matchingParen(trimmed, 0);
    // An unbalanced / non-terminal `(` is a malformed expression, NOT a literal — surface a
    // corrective so the model fixes it, rather than silently writing the raw `(...` string.
    if (close === -1 || close !== trimmed.length - 1) {
      return { error: 'unbalanced ( ) in expression — wrap exactly one pipeline: ( <pipeline> )' };
    }
    const inner = trimmed.slice(1, close).trim();
    if (inner === '') return { error: 'empty ( ) expression — put a pipeline inside the parens' };
    return parsePipeline(inner);
  }
  return undefined; // a literal — caller keeps it verbatim (ADR-0004 back-compat)
}

/**
 * Index of the `)` matching the `(` at `open`, honoring nested parens but ignoring parens inside a
 * double-quoted span (so a quoted pipeline arg with a `)` doesn't mis-close). Returns -1 when there
 * is no matching close.
 */
function matchingParen(s: string, open: number): number {
  let depth = 0;
  let i = open;
  while (i < s.length) {
    const ch = s[i]!;
    if (ch === '"') {
      const quoted = scanQuoted(s, i);
      if (quoted) {
        i = quoted.end;
        continue;
      }
      return -1; // unterminated quote
    }
    if (ch === '(') depth++;
    else if (ch === ')') {
      depth--;
      if (depth === 0) return i;
    }
    i++;
  }
  return -1;
}

/* ───────────────────────────── parsing ───────────────────────────── */

/** A parse error carries a CLI-style corrective message (same shape as the command parser's). */
export interface ExprParseError {
  error: string;
}

export function isExprParseError(node: ParsedExpr | ExprParseError): node is ExprParseError {
  return 'error' in node;
}

/**
 * Parse one expression line into a `ParsedExpr` or a corrective `{ error }`.
 *
 *   let $name = <pipeline>          → bind the pipeline's value to $name
 *   <source> ( '|' <stage> )*       → a pipeline
 *
 * The caller (`parseProgramBlock`) routes only expression lines here; a non-expression line is
 * parsed by the ADR-0004 command parser. Defensive: a malformed expression yields a corrective
 * error, never a throw.
 */
export function parseExpressionLine(line: string): ParsedExpr | ExprParseError {
  const trimmed = line.trim();

  if (/^let\s/i.test(trimmed)) {
    return parseLet(trimmed);
  }
  return parsePipeline(trimmed);
}

/** `let $name = <pipeline>` — a `$`-prefixed identifier, an `=`, then a pipeline. */
function parseLet(line: string): ParsedExpr | ExprParseError {
  const usage = 'let needs a $name and a pipeline — usage: let $name = <source> | <transform> ...';
  // Strip the leading `let`.
  const afterLet = line.slice(line.search(/\s/) + 1).trim();
  const eq = afterLet.indexOf('=');
  if (eq === -1) return { error: usage };

  const name = afterLet.slice(0, eq).trim();
  const rhs = afterLet.slice(eq + 1).trim();
  if (!isVarName(name)) {
    return { error: `let needs a $-prefixed name (e.g. let $t = ...) — got "${name}"` };
  }
  if (rhs === '') return { error: usage };

  const pipeline = parsePipeline(rhs);
  if (isExprParseError(pipeline)) return pipeline;
  return { kind: 'let', name: stripDollar(name), pipeline };
}

/** `<source> ( '|' <stage> )*` — split on top-level pipes, parse the head as a source. */
function parsePipeline(text: string): PipelineExpr | ExprParseError {
  const segments = splitTopLevelPipes(text);
  const head = segments[0]!.trim();
  if (head === '') return { error: 'pipeline needs a source — e.g. read <selector> | <transform>' };

  const source = parseSource(head);
  if ('error' in source) return source;

  const stages: TransformStage[] = [];
  for (let i = 1; i < segments.length; i++) {
    const seg = segments[i]!.trim();
    if (seg === '') return { error: 'empty pipeline stage (a `|` with nothing after it)' };
    const sp = seg.search(/\s/);
    // Transform names are case-insensitive (`| SUM x` ≡ `| sum x`), matching the Python preflight;
    // the runtime registry is keyed by the lowercase name.
    const name = (sp === -1 ? seg : seg.slice(0, sp)).toLowerCase();
    const args = sp === -1 ? '' : seg.slice(sp + 1).trim();
    stages.push({ name, args });
  }
  return { kind: 'pipeline', source, stages };
}

/** The largest literal table a program line may carry (rows × columns), so a pasted blob can't bloat a turn. */
export const MAX_LITERAL_TABLE_CELLS = 2000;

/** A numeric literal source: an optionally signed integer or decimal (`5`, `-3.25`). */
const NUMBER_LITERAL_RE = /^-?\d+(?:\.\d+)?$/;

/**
 * A pipeline source: `read <selector>`, `search <text>`, `outline`, `$var`, or a C1 literal
 * (`<number>`, `"<text>"`, `table [[…], …]`). `table` is also an Excel EFFECT verb (promote a range
 * to a native table); in source position it is a literal ONLY when followed by a `[` — any other
 * `table …` head is the effect and stays refused by the pure-only guard.
 */
function parseSource(text: string): PipeSource | ExprParseError {
  if (isVarName(text)) return { src: 'var', name: stripDollar(text) };
  if (NUMBER_LITERAL_RE.test(text)) return { src: 'number', value: Number(text) };
  if (text.startsWith('"')) {
    const quoted = scanQuoted(text, 0);
    if (!quoted || text.slice(quoted.end).trim() !== '') {
      return {
        error: 'text literal source must be ONE double-quoted string — e.g. let $s = "hello"',
      };
    }
    return { src: 'text', value: quoted.value };
  }

  const sp = text.search(/\s/);
  const head = (sp === -1 ? text : text.slice(0, sp)).toLowerCase();
  const rest = sp === -1 ? '' : text.slice(sp + 1).trim();

  switch (head) {
    case 'read':
      // Empty selector is valid (whole-document read, as in ADR-0004).
      return { src: 'read', selector: rest };
    case 'search':
      if (rest === '') return { error: 'search source needs text — e.g. search revenue' };
      return { src: 'search', text: stripWrappingQuotes(rest) };
    case 'outline':
      return { src: 'outline' };
    case 'table':
      if (rest.startsWith('[')) return parseTableLiteral(rest);
      return { error: EFFECT_COMPOSE_ERROR };
    default:
      // A pipe into / out of an effect verb (`set X | …`, `read X | set …`) — Phase 1 is pure-only.
      if (EFFECT_VERBS.has(head)) return { error: EFFECT_COMPOSE_ERROR };
      return {
        error: `unknown pipeline source "${head}" — use read <selector>, search <text>, outline, $var, a number, a "text" literal, or table [[…]]`,
      };
  }
}

/**
 * `table [["col","col"], ["cell","cell"], …]` — a JSON array of rows; the first row is the header.
 * Cells may be strings, numbers, booleans or null (stringified; null → ""). Rows are padded or
 * truncated to the header width so the value is rectangular, exactly like a parsed GFM read.
 */
function parseTableLiteral(json: string): PipeSource | ExprParseError {
  const usage =
    'table literal usage: table [["col1","col2"],["a","1"],["b","2"]] (first row = header)';
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    return { error: `table literal is not valid JSON — ${usage}` };
  }
  if (!Array.isArray(parsed) || parsed.length === 0 || !parsed.every(Array.isArray)) {
    return { error: `table literal must be a non-empty array of row arrays — ${usage}` };
  }
  const rowsRaw = parsed as unknown[][];
  const toCell = (c: unknown): string | undefined => {
    if (c === null || c === undefined) return '';
    if (typeof c === 'string') return c;
    if (typeof c === 'number' || typeof c === 'boolean') return String(c);
    return undefined;
  };
  const header = rowsRaw[0]!.map(toCell);
  if (header.length === 0 || header.some((c) => c === undefined || c.trim() === '')) {
    return { error: `table literal header row needs non-empty column names — ${usage}` };
  }
  const columns = header.map((c) => c!.trim());
  if (new Set(columns).size !== columns.length) {
    return { error: 'table literal header has duplicate column names' };
  }
  if (rowsRaw.length * columns.length > MAX_LITERAL_TABLE_CELLS) {
    return {
      error: `table literal too large (max ${MAX_LITERAL_TABLE_CELLS} cells) — read it from the document instead`,
    };
  }
  const rows: string[][] = [];
  for (let r = 1; r < rowsRaw.length; r++) {
    const cells = rowsRaw[r]!.map(toCell);
    if (cells.some((c) => c === undefined)) {
      return { error: `table literal row ${r + 1} has a non-scalar cell — ${usage}` };
    }
    const row = (cells as string[]).slice(0, columns.length);
    while (row.length < columns.length) row.push('');
    rows.push(row);
  }
  return { src: 'table', columns, rows };
}

/* ───────────────────────── top-level pipe splitting ─────────────────────── */

/**
 * Split a line on top-level `|` characters — pipes NOT inside a double-quoted span (so a quoted
 * search/comment arg containing a literal `|` is preserved). Honors `\"` escapes via the shared
 * scanner. Returns the segments between pipes (without the pipe characters).
 */
function splitTopLevelPipes(text: string): string[] {
  const cuts = findTopLevelPipes(text);
  if (cuts.length === 0) return [text];
  const segments: string[] = [];
  let prev = 0;
  for (const cut of cuts) {
    segments.push(text.slice(prev, cut));
    prev = cut + 1;
  }
  segments.push(text.slice(prev));
  return segments;
}

/**
 * The indices of every top-level (unquoted, unparenthesized) `|` in `text`. Pipes inside a
 * double-quoted span (a quoted comment body) and inside a parenthesized span (an ADR-0005 Phase-2
 * effect-arg expression, `set X = ($a | sum Y)`) are NOT top-level — so such a `set`/`comment` line
 * is NOT classified as a standalone pipeline and routes to the ADR-0004 command parser, which then
 * parses the parenthesized arg as an effect-arg expression.
 */
function findTopLevelPipes(text: string): number[] {
  const cuts: number[] = [];
  let depth = 0;
  let i = 0;
  while (i < text.length) {
    const ch = text[i]!;
    if (ch === '"') {
      const quoted = scanQuoted(text, i);
      if (quoted) {
        i = quoted.end;
        continue;
      }
      // Unterminated quote: stop scanning for pipes inside it (treat the rest as literal).
      break;
    }
    if (ch === '(') depth++;
    else if (ch === ')' && depth > 0) depth--;
    else if (ch === '|' && depth === 0) cuts.push(i);
    i++;
  }
  return cuts;
}

/* ───────────────────────────── helpers ───────────────────────────── */

/** A `$`-prefixed identifier (`$t`, `$total_q1`). */
function isVarName(s: string): boolean {
  return /^\$[A-Za-z_][A-Za-z0-9_]*$/.test(s);
}

function stripDollar(s: string): string {
  return s.startsWith('$') ? s.slice(1) : s;
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
