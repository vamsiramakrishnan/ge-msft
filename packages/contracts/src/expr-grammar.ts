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
