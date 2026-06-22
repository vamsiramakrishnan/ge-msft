import { z } from 'zod';
import { scanQuoted } from './command-grammar.js';

/**
 * The ADR-0004 effect verbs — they cannot be a pipeline source in Phase 1 (pure-only). Inlined
 * (not imported from `WRITE_VERB_TO_KIND`) to avoid a module-eval-time dependency on the
 * command-grammar module (the two modules reference each other), and kept in sync by the
 * exhaustiveness test in `expr-grammar.test.ts`.
 */
export const EFFECT_VERBS: ReadonlySet<string> = new Set(['set', 'suggest', 'comment', 'format']);

/** The Phase-1 corrective when the model tries to compose an effect. */
export const EFFECT_COMPOSE_ERROR =
  "effects can't be composed yet — emit them as standalone commands (Phase 2)";

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
 *   <source> = `read <selector>` | `search <text>` | `outline` | `$var`.
 */

/* ───────────────────────────── AST ───────────────────────────── */

/** The head of a pipeline — where the value comes from before any transform runs. */
export type PipeSource =
  | { src: 'read'; selector: string }
  | { src: 'search'; text: string }
  | { src: 'outline' }
  | { src: 'var'; name: string };

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
    const name = sp === -1 ? seg : seg.slice(0, sp);
    const args = sp === -1 ? '' : seg.slice(sp + 1).trim();
    stages.push({ name, args });
  }
  return { kind: 'pipeline', source, stages };
}

/** A pipeline source: `read <selector>`, `search <text>`, `outline`, or `$var`. */
function parseSource(text: string): PipeSource | ExprParseError {
  if (isVarName(text)) return { src: 'var', name: stripDollar(text) };

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
    default:
      // A pipe into / out of an effect verb (`set X | …`, `read X | set …`) — Phase 1 is pure-only.
      if (EFFECT_VERBS.has(head)) return { error: EFFECT_COMPOSE_ERROR };
      return {
        error: `unknown pipeline source "${head}" — use read <selector>, search <text>, outline, or $var`,
      };
  }
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

/** The indices of every top-level (unquoted) `|` in `text`. */
function findTopLevelPipes(text: string): number[] {
  const cuts: number[] = [];
  let i = 0;
  while (i < text.length) {
    if (text[i] === '"') {
      const quoted = scanQuoted(text, i);
      if (quoted) {
        i = quoted.end;
        continue;
      }
      // Unterminated quote: stop scanning for pipes inside it (treat the rest as literal).
      break;
    }
    if (text[i] === '|') cuts.push(i);
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
