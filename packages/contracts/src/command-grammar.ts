import { z } from 'zod';
import { isAnalysisBindingKind } from './analysis-actions.js';
import { ActuationKindSchema, type ActuationKind, type CapabilityManifest } from './capability.js';
import {
  ParsedExprSchema,
  isExpressionLine,
  isExprParseError,
  parseEffectArg,
  parseExpressionLine,
  type ExprParseError,
  type ParsedExpr,
} from './expr-grammar.js';
import {
  isSkillDefHeader,
  isSkillEnd,
  parseSkillCall,
  parseSkillDefHeader,
  type ParsedSkillCall,
  type ParsedSkillDef,
} from './skill-grammar.js';
import { PlanContextHintSchema, type PlanContextHint } from './command-plan.js';
import { ContextKindSchema, type ContextKind, type Surface } from './context.js';
import { registryEntryForKindAndSurface } from './capability-registry.js';

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
 * SCOPE: `outline · read · search · list · inspect · properties · comments · attachments · tables ·
 * slides · neighbors · context · open · workspace · save · cat · grep · cp · mv · rm · share · set ·
 * grid · suggest · comment · format · reply · slide · page · mail · post · compose · table · chart · cf · spill ·
 * done · help` (ADR-0007 adds the host-native kinds; the workspace verbs are local/non-host
 * artifact operations).
 */

/**
 * Read verbs. `outline`/`read`/`search` are Layer-B host reads (ADR-0003). `context` is a
 * runtime-served read-only strategy probe: it never uploads a file, runs code, or writes content.
 */
export const READ_VERBS = [
  'outline',
  'read',
  'search',
  'ls',
  'find',
  'tail',
  'list',
  'inspect',
  'properties',
  'comments',
  'attachments',
  'tables',
  'slides',
  'neighbors',
  'context',
  'open',
] as const;

/** Control verbs. Always advertised; not actuations. */
export const CONTROL_VERBS = ['done', 'help'] as const;

/**
 * Workspace verbs are local, non-host operations over bounded virtual artifacts. They can read from
 * host reads or pure composed values, but they never mutate Office content and never bypass the
 * write approval gate. The runtime stores only bounded data and returns compact artifact handles.
 *
 * `share` is the one exception to "local": it has the exact same source grammar as `save`, but
 * persists to the cross-surface `/shared` handoff store (Microsoft Graph app-folder, see
 * `@ge/graph-client`'s `GRAPH_SCOPES.shared`) instead of the in-memory `/work` store — so a value
 * saved from Excel can be read back by name from Word/PowerPoint/Teams. It still never touches
 * Office content and never bypasses the write approval gate; it is gated by whether the runtime was
 * given a `sharedStore` for this session at all (Graph consent may not be granted).
 */
export const WORKSPACE_VERBS = [
  'analyze',
  'workspace',
  'save',
  'cat',
  'grep',
  'cp',
  'mv',
  'rm',
  'share',
] as const;

/**
 * Write verbs → the `ActuationKind` (ADR-0002) they compile to. A write verb is advertised
 * for a surface ONLY when the surface's `CapabilityManifest.actuations[]` includes its mapped
 * kind (so `set` shows for Excel's `write-cells`, `suggest` for Word's `tracked-change`). The
 * parser, the advertisement, and the runtime compiler all derive from this single map.
 */
export const WRITE_VERB_TO_KIND = {
  set: 'write-cells',
  grid: 'write-cells',
  suggest: 'tracked-change',
  comment: 'add-comment',
  format: 'format-cells',
  reply: 'comment-reply',
  // ADR-0006 CLI parity — the bridges already HANDLE + advertise these kinds; these verbs make
  // them reachable from the model. Each compiles to the kind the matching bridge's actuate()
  // consumes (PowerPoint `insert-slide`, OneNote `append-page`, Outlook `reply-mail`, Teams
  // `post-message`), and each is gated/approved + provenanced like every other effect.
  slide: 'insert-slide',
  page: 'append-page',
  mail: 'reply-mail',
  post: 'post-message',
  // Outlook `create-mail` — compose a brand-new grounded draft (vs `mail`, which replies to the open
  // item). The bridge opens a fresh message form via displayNewMessageForm; recipients are left for
  // the user to fill (we never auto-address). Gated/approved + provenanced like every other effect.
  compose: 'create-mail',
  // ADR-0007 host-native Excel write kinds — each is gated/approved + provenanced and carries a
  // recorded inverse (delete-object / restore / clear-rule). `table`/`chart`/`cf` take a literal
  // range + props (the `format`-style grammar); the data source for a chart is a range produced
  // directly or by composition (`spill`, ADR-0007 §3).
  table: 'create-table',
  chart: 'insert-chart',
  cf: 'format-conditional',
  // PowerPoint shape/textbox surgery: replace the text of one explicitly addressed shape.
  shape: 'set-shape-text',
  // `spill` is `write-cells` widened (ADR-0007 §3): its arg is a TABLE expression whose rows are
  // written as a grid (the missing table→cells sink), vs `set` which writes one scalar. It reuses
  // the existing write-cells kind/bridge/safety — no new host work — so it advertises wherever
  // write-cells does. Many verbs → one kind is fine (the map is verb→kind).
  spill: 'write-cells',
} satisfies Record<string, ActuationKind>;

export type WriteVerb = keyof typeof WRITE_VERB_TO_KIND;
export type ReadVerb = (typeof READ_VERBS)[number];
export type ControlVerb = (typeof CONTROL_VERBS)[number];
export type WorkspaceVerb = (typeof WORKSPACE_VERBS)[number];

/** Every verb this slice understands (for did-you-mean + advertisement). */
export const ALL_VERBS = [
  ...READ_VERBS,
  ...WORKSPACE_VERBS,
  ...CONTROL_VERBS,
  ...(Object.keys(WRITE_VERB_TO_KIND) as WriteVerb[]),
] as const;

/* ───────────────────────────── ParsedCommand ───────────────────────────── */

/**
 * A successfully parsed command line — a discriminated union on `verb`. The selector/args
 * are surface-agnostic strings here; the runtime compiler maps them onto surface-specific
 * targets (`set` cell → `target.range`; `suggest` oldText → `target.matchText`).
 *
 * ADR-0005 Phase 2 — the effect verbs carry an optional `*Expr`: a `ParsedExpr` the runtime
 * evaluates against the binding env at dry-run, rendering the resulting `Value` into the concrete
 * param. Composition parity (this wave): every text-bearing effect verb is expression-bearing —
 * `set`→`valueExpr`, `comment`/`reply`→`textExpr`, `mail`/`page`/`compose`→`bodyExpr`,
 * `post`→`textExpr`, and `slide`→`bulletsExpr` (a table whose rows become bullets). When the arg is
 * a LITERAL (plain text or `=formula`, the ADR-0004 case) the `*Expr` field is absent and the param
 * carries the verbatim literal. `suggest` and `format` (typed key=value props, no free-text slot)
 * stay literal-only.
 */
export type ParsedCommand =
  | { verb: 'outline' }
  | { verb: 'read'; selector: string }
  | { verb: 'search'; text: string }
  | { verb: 'ls'; path: string }
  | { verb: 'find'; path: string; glob?: string }
  | { verb: 'tail'; path: string; n?: number }
  | { verb: 'list'; kind?: ContextKind }
  | { verb: 'inspect'; selector: string }
  | { verb: 'properties'; selector: string }
  | { verb: 'comments'; selector?: string }
  | { verb: 'attachments'; selector?: string }
  | { verb: 'tables'; selector?: string }
  | { verb: 'slides'; selector?: string }
  | { verb: 'neighbors'; selector?: string }
  | { verb: 'context'; hints: PlanContextHint[] }
  | { verb: 'open'; selector: string }
  | { verb: 'analyze'; request: string }
  | { verb: 'workspace'; ref?: string }
  | { verb: 'save'; name: string; source: WorkspaceSource }
  | { verb: 'cat'; ref: string; head?: number }
  | { verb: 'grep'; ref: string; pattern: string; context?: number }
  | { verb: 'cp'; src: string; dst: string }
  | { verb: 'mv'; src: string; dst: string }
  | { verb: 'rm'; name: string }
  | { verb: 'share'; name: string; source: WorkspaceSource }
  | { verb: 'set'; cell: string; value: string; valueExpr?: ParsedExpr }
  | { verb: 'grid'; range: string; cells: string[][] }
  | { verb: 'suggest'; oldText: string; newText: string }
  | { verb: 'comment'; selector: string; text: string; textExpr?: ParsedExpr }
  | { verb: 'format'; range: string; props: Record<string, string> }
  | { verb: 'reply'; commentId: string; text: string; textExpr?: ParsedExpr }
  // ADR-0006 CLI parity effect verbs — now composition-bearing: the free-text slot accepts a
  // `( <pipeline> )` / `$var` expression (evaluated at dry-run) as well as a quoted literal. The
  // anchor slots (title/subject) stay literal. `slide` bullets accept a table expression whose rows
  // become bullets (`bulletsExpr`).
  | { verb: 'slide'; title: string; bullets: string[]; bulletsExpr?: ParsedExpr }
  | { verb: 'page'; title: string; body: string; bodyExpr?: ParsedExpr }
  | { verb: 'mail'; body: string; bodyExpr?: ParsedExpr }
  | { verb: 'post'; text: string; textExpr?: ParsedExpr }
  | { verb: 'compose'; subject: string; body: string; bodyExpr?: ParsedExpr }
  // ADR-0007 host-native Excel kinds — literal range + props (no free-text/expression slot in this
  // wave; the chart title is a literal). The anchor is always an explicit range/name.
  | { verb: 'table'; range: string; props: Record<string, string> }
  | { verb: 'chart'; chartType: string; range: string; props: Record<string, string> }
  | { verb: 'cf'; range: string; props: Record<string, string> }
  | { verb: 'shape'; selector: string; text: string; textExpr?: ParsedExpr }
  // ADR-0007 §3 — the table→grid composition sink. `valueExpr` is the source TABLE expression
  // (pre-resolution); `cells` is the resolved grid (filled by the runtime at dry-run). Spill is
  // ALWAYS expression-driven — a bare literal grid is not expressible inline.
  | { verb: 'spill'; range: string; valueExpr?: ParsedExpr; cells?: string[][] }
  // ADR-0008 §two-tier — the `/<kind>` SPECIALIZED surface. A named, non-composing effect terminal
  // for the long-tail catalogue (`/insert-image base64=… alt="…"`). `kind` is the `ActuationKind`
  // itself (no alias table → drift-free); `props`/`args` are the typed/positional arguments the
  // runtime compiles into the kind's `ActuationParams`. Per-surface availability is checked at
  // compile, not here (the parser stays structural).
  | { verb: 'invoke'; kind: string; props: Record<string, string>; args: string[] }
  | { verb: 'done' }
  | { verb: 'help'; topic?: string };

export type WorkspaceSource =
  | { src: 'outline' }
  | { src: 'read'; selector: string }
  | { src: 'search'; text: string }
  | { src: 'literal'; text: string }
  | { src: 'expr'; expr: ParsedExpr };

export const ParsedCommandSchema: z.ZodType<ParsedCommand> = z.discriminatedUnion('verb', [
  z.object({ verb: z.literal('outline') }),
  z.object({ verb: z.literal('read'), selector: z.string() }),
  z.object({ verb: z.literal('search'), text: z.string() }),
  z.object({ verb: z.literal('ls'), path: z.string() }),
  z.object({ verb: z.literal('find'), path: z.string(), glob: z.string().optional() }),
  z.object({ verb: z.literal('tail'), path: z.string(), n: z.number().optional() }),
  z.object({ verb: z.literal('list'), kind: ContextKindSchema.optional() }),
  z.object({ verb: z.literal('inspect'), selector: z.string() }),
  z.object({ verb: z.literal('properties'), selector: z.string() }),
  z.object({ verb: z.literal('comments'), selector: z.string().optional() }),
  z.object({ verb: z.literal('attachments'), selector: z.string().optional() }),
  z.object({ verb: z.literal('tables'), selector: z.string().optional() }),
  z.object({ verb: z.literal('slides'), selector: z.string().optional() }),
  z.object({ verb: z.literal('neighbors'), selector: z.string().optional() }),
  z.object({ verb: z.literal('context'), hints: z.array(PlanContextHintSchema) }),
  z.object({ verb: z.literal('open'), selector: z.string() }),
  z.object({ verb: z.literal('analyze'), request: z.string().max(32768) }),
  z.object({ verb: z.literal('workspace'), ref: z.string().optional() }),
  z.object({
    verb: z.literal('save'),
    name: z.string(),
    source: z.union([
      z.object({ src: z.literal('outline') }),
      z.object({ src: z.literal('read'), selector: z.string() }),
      z.object({ src: z.literal('search'), text: z.string() }),
      z.object({ src: z.literal('literal'), text: z.string() }),
      z.object({ src: z.literal('expr'), expr: z.lazy(() => ParsedExprSchema) }),
    ]),
  }),
  z.object({
    verb: z.literal('cat'),
    ref: z.string(),
    head: z.number().int().positive().optional(),
  }),
  z.object({
    verb: z.literal('grep'),
    ref: z.string(),
    pattern: z.string(),
    context: z.number().int().min(0).optional(),
  }),
  z.object({ verb: z.literal('cp'), src: z.string(), dst: z.string() }),
  z.object({ verb: z.literal('mv'), src: z.string(), dst: z.string() }),
  z.object({ verb: z.literal('rm'), name: z.string() }),
  z.object({
    verb: z.literal('share'),
    name: z.string(),
    source: z.union([
      z.object({ src: z.literal('outline') }),
      z.object({ src: z.literal('read'), selector: z.string() }),
      z.object({ src: z.literal('search'), text: z.string() }),
      z.object({ src: z.literal('literal'), text: z.string() }),
      z.object({ src: z.literal('expr'), expr: z.lazy(() => ParsedExprSchema) }),
    ]),
  }),
  z.object({
    verb: z.literal('set'),
    cell: z.string(),
    value: z.string(),
    valueExpr: z.lazy(() => ParsedExprSchema).optional(),
  }),
  z.object({ verb: z.literal('grid'), range: z.string(), cells: z.array(z.array(z.string())) }),
  z.object({ verb: z.literal('suggest'), oldText: z.string(), newText: z.string() }),
  z.object({
    verb: z.literal('comment'),
    selector: z.string(),
    text: z.string(),
    textExpr: z.lazy(() => ParsedExprSchema).optional(),
  }),
  z.object({ verb: z.literal('format'), range: z.string(), props: z.record(z.string()) }),
  z.object({
    verb: z.literal('reply'),
    commentId: z.string(),
    text: z.string(),
    textExpr: z.lazy(() => ParsedExprSchema).optional(),
  }),
  z.object({
    verb: z.literal('slide'),
    title: z.string(),
    bullets: z.array(z.string()),
    bulletsExpr: z.lazy(() => ParsedExprSchema).optional(),
  }),
  z.object({
    verb: z.literal('page'),
    title: z.string(),
    body: z.string(),
    bodyExpr: z.lazy(() => ParsedExprSchema).optional(),
  }),
  z.object({
    verb: z.literal('mail'),
    body: z.string(),
    bodyExpr: z.lazy(() => ParsedExprSchema).optional(),
  }),
  z.object({
    verb: z.literal('post'),
    text: z.string(),
    textExpr: z.lazy(() => ParsedExprSchema).optional(),
  }),
  z.object({
    verb: z.literal('compose'),
    subject: z.string(),
    body: z.string(),
    bodyExpr: z.lazy(() => ParsedExprSchema).optional(),
  }),
  z.object({ verb: z.literal('table'), range: z.string(), props: z.record(z.string()) }),
  z.object({
    verb: z.literal('chart'),
    chartType: z.string(),
    range: z.string(),
    props: z.record(z.string()),
  }),
  z.object({ verb: z.literal('cf'), range: z.string(), props: z.record(z.string()) }),
  z.object({
    verb: z.literal('shape'),
    selector: z.string(),
    text: z.string(),
    textExpr: z.lazy(() => ParsedExprSchema).optional(),
  }),
  z.object({
    verb: z.literal('spill'),
    range: z.string(),
    valueExpr: z.lazy(() => ParsedExprSchema).optional(),
    cells: z.array(z.array(z.string())).optional(),
  }),
  z.object({
    verb: z.literal('invoke'),
    kind: z.string(),
    props: z.record(z.string()),
    args: z.array(z.string()),
  }),
  z.object({ verb: z.literal('done') }),
  z.object({ verb: z.literal('help'), topic: z.string().optional() }),
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
  const match = modelText.match(/^```cmd[^\S\n]*\r?\n([\s\S]*?)^```[^\S\n]*(?:\r?\n|$)/im);
  if (match) return match[1]!.trim();

  // Live StreamAssist can occasionally stream the opening cmd fence and finish the answer without
  // the closing fence. Recover only the whole-response shape: the first non-empty bytes must be the
  // cmd fence, and no later fence marker may appear. Any prose before the fence still fails closed.
  const unclosed = modelText.trim();
  const unclosedMatch = /^```cmd[^\S\n]*\r?\n([\s\S]+)$/i.exec(unclosed);
  if (unclosedMatch && !unclosedMatch[1]!.includes('```')) {
    return unclosedMatch[1]!.trim();
  }

  // Some StreamAssist surfaces have returned the language marker as a plain first line (`cmd`)
  // instead of preserving the Markdown fence. Accept only that whole-response shape: first
  // non-empty line exactly `cmd`, followed by command lines. Prose before `cmd` still fails closed
  // as a no-fence turn, and other plain markers (`python`, `json`, `bash`) remain non-executable.
  const plain = modelText.trim();
  const plainMatch = /^cmd[^\S\n]*\r?\n([\s\S]+)$/i.exec(plain);
  return plainMatch ? plainMatch[1]!.trim() : null;
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
 *   • `context` — read-only context/upload/code-exec strategy hints (validated enum tokens).
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

  // ADR-0008 §two-tier — a `/<kind>` line is the SPECIALIZED surface (the long-tail catalogue),
  // dispatched before the core-verb switch. The bare verbs stay the composable algebra.
  if (rest === '-h' || rest === '--help') return { verb: 'help', topic: rawVerb };
  if (rawVerb.startsWith('/')) return parseInvoke(rawVerb, rest);

  switch (verb) {
    case 'outline':
      return { verb: 'outline' };
    case 'done':
      return { verb: 'done' };
    case 'help':
      return { verb: 'help', ...(rest ? { topic: rest } : {}) };

    case 'read': {
      // Excel: an A1/NamedRange selector. Word: whole-doc (no selector) — empty string is fine,
      // the runtime read-intent treats an empty selector as "whole document".
      return { verb: 'read', selector: rest };
    }

    case 'ls': {
      if (rest === '') return { error: 'ls needs a path — usage: ls <path>, e.g. ls /doc' };
      return { verb: 'ls', path: rest };
    }

    case 'find': {
      if (rest === '') return { error: 'find needs a path — usage: find <path> [glob]' };
      const [path, glob] = rest.split(/\s+/, 2);
      return { verb: 'find', path: path!, ...(glob ? { glob } : {}) };
    }

    case 'tail': {
      if (rest === '') return { error: 'tail needs a path — usage: tail <path> [n]' };
      const [path, nStr] = rest.split(/\s+/, 2);
      const n = nStr ? Number(nStr) : undefined;
      if (nStr !== undefined && (n === undefined || Number.isNaN(n))) {
        return { error: 'tail: n must be a number' };
      }
      return { verb: 'tail', path: path!, ...(n !== undefined ? { n } : {}) };
    }

    case 'search': {
      if (rest === '') return { error: 'search needs text — usage: search <text>' };
      // Tolerate the model wrapping the query in quotes.
      return { verb: 'search', text: stripWrappingQuotes(rest) };
    }

    case 'list': {
      if (rest === '') return { verb: 'list' };
      const kind = rest.toLowerCase();
      const parsed = ContextKindSchema.safeParse(kind);
      if (!parsed.success) {
        return {
          error: `unknown context kind "${rest}" — supported: ${ContextKindSchema.options.join(', ')}`,
        };
      }
      return { verb: 'list', kind: parsed.data };
    }

    case 'inspect': {
      if (rest === '')
        return { error: 'inspect needs a ref id or selector — usage: inspect <refId|selector>' };
      return { verb: 'inspect', selector: stripWrappingQuotes(rest) };
    }

    case 'properties': {
      if (rest === '')
        return {
          error: 'properties needs a ref id or selector — usage: properties <refId|selector>',
        };
      return { verb: 'properties', selector: stripWrappingQuotes(rest) };
    }

    case 'comments':
      return { verb: 'comments', ...(rest ? { selector: stripWrappingQuotes(rest) } : {}) };
    case 'attachments':
      return { verb: 'attachments', ...(rest ? { selector: stripWrappingQuotes(rest) } : {}) };
    case 'tables':
      return { verb: 'tables', ...(rest ? { selector: stripWrappingQuotes(rest) } : {}) };
    case 'slides':
      return { verb: 'slides', ...(rest ? { selector: stripWrappingQuotes(rest) } : {}) };
    case 'neighbors':
      return { verb: 'neighbors', ...(rest ? { selector: stripWrappingQuotes(rest) } : {}) };

    case 'context': {
      if (rest === '') return { verb: 'context', hints: [] };
      const hints: PlanContextHint[] = [];
      for (const raw of rest.split(/\s+/).filter(Boolean)) {
        const hint = raw.toLowerCase();
        const parsed = PlanContextHintSchema.safeParse(hint);
        if (!parsed.success) {
          return {
            error: `unknown context hint "${raw}" — supported: ${PlanContextHintSchema.options.join(', ')}`,
          };
        }
        hints.push(parsed.data);
      }
      return { verb: 'context', hints };
    }

    case 'open': {
      if (rest === '')
        return { error: 'open needs a ref id or selector — usage: open <refId|selector>' };
      return { verb: 'open', selector: stripWrappingQuotes(rest) };
    }

    case 'analyze':
      if (!rest || rest.length > 32768)
        return { error: 'analyze requires a bounded JSON action object' };
      try {
        const value: unknown = JSON.parse(rest);
        if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error();
      } catch {
        return { error: 'analyze requires a JSON action object' };
      }
      return { verb: 'analyze', request: rest };
    case 'workspace':
      return { verb: 'workspace', ...(rest ? { ref: stripWrappingQuotes(rest) } : {}) };

    case 'save':
      return parseSave(rest);

    case 'cat':
      return parseCat(rest);

    case 'grep':
      return parseGrep(rest);

    case 'cp':
      return parseCp(rest);

    case 'mv':
      return parseMv(rest);

    case 'rm':
      return parseRm(rest);

    case 'share':
      return parseShare(rest);

    case 'set': {
      const split = splitFirstArg(rest);
      if (!split) {
        return { error: 'set needs a cell and a value — usage: set <A1> <value|=formula>' };
      }
      const cell = split.arg;
      const value = unquoteSetValue(split.tail.trim());
      if (cell === '' || value === '') {
        return { error: 'set needs a cell and a value — usage: set <A1> <value|=formula>' };
      }
      // ADR-0005 Phase 2: the value may be an effect-arg EXPRESSION (`$var` or `( <pipeline> )`)
      // evaluated against the env at dry-run; otherwise it is a LITERAL (commas/spaces/`=SUM(A1,
      // A2)` preserved verbatim, ADR-0004 back-compat).
      const expr = parseEffectArg(value);
      if (expr === undefined) return { verb: 'set', cell, value };
      if (isExprParseError(expr)) return expr;
      return { verb: 'set', cell, value, valueExpr: expr };
    }

    case 'grid':
      return parseGrid(rest);

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

    case 'slide':
      return parseSlide(rest);
    case 'page':
      return parsePage(rest);
    case 'mail':
      return parseMail(rest);
    case 'post':
      return parsePost(rest);
    case 'compose':
      return parseCompose(rest);

    case 'table':
      return parseTable(rest);
    case 'chart':
      return parseChart(rest);
    case 'cf':
      return parseCf(rest);
    case 'shape':
      return parseShape(rest);
    case 'spill':
      return parseSpill(rest);

    default:
      return { error: unknownVerbError(verb) };
  }
}

/**
 * `grid <range> = "a\tb\nc\td"` — write a rectangular literal grid as ONE write-cells effect.
 * This is the bulk-write sibling of `set`: the quoted body is TSV with escaped `\t` and `\n`.
 */
function parseGrid(rest: string): ParsedCommand | CommandParseError {
  const usage = 'grid needs a range and quoted TSV — usage: grid <range> = "a\\tb\\nc\\td"';
  const split = splitFirstArg(rest);
  if (!split) return { error: usage };
  const range = split.arg;
  let tail = split.tail.trim();
  if (tail.startsWith('=')) tail = tail.slice(1).trim();
  if (range === '' || tail === '') return { error: usage };
  const quoted = scanQuoted(tail, 0);
  if (!quoted || tail.slice(quoted.end).trim() !== '') return { error: usage };
  const cells = parseGridBody(quoted.value);
  if (cells.length === 0 || cells.every((row) => row.every((cell) => cell === ''))) {
    return { error: 'grid body is empty — provide at least one non-empty cell' };
  }
  const width = cells[0]?.length ?? 0;
  if (width === 0 || cells.some((row) => row.length !== width)) {
    return { error: 'grid rows must be rectangular — every row needs the same number of cells' };
  }
  return { verb: 'grid', range, cells };
}

function parseGridBody(body: string): string[][] {
  const normalized = body.replace(/\\n/g, '\n').replace(/\\t/g, '\t');
  return normalized
    .split(/\r?\n/)
    .map((row) => row.split('\t').map((cell) => cell.trim()))
    .filter((row) => row.some((cell) => cell !== ''));
}

/**
 * `save <name> = read <selector>` stores a bounded local artifact without writing Office content.
 * Sources may be `outline`, `read`, `search`, a quoted literal, or a pure composition expression.
 */
function parseSave(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'save needs a name and source — usage: save <name> = read <selector> | search <text> | outline | "literal" | ($pipeline)';
  const split = splitFirstArg(rest);
  if (!split) return { error: usage };
  const name = normalizeWorkspaceName(split.arg);
  if (typeof name !== 'string') return name;
  let tail = split.tail.trim();
  if (tail.startsWith('=')) tail = tail.slice(1).trim();
  if (tail === '') return { error: usage };

  if (tail.toLowerCase() === 'outline') return { verb: 'save', name, source: { src: 'outline' } };
  if (/^read(?:\s|$)/i.test(tail)) {
    return {
      verb: 'save',
      name,
      source: { src: 'read', selector: tail.replace(/^read\s*/i, '').trim() },
    };
  }
  if (/^search\s+/i.test(tail)) {
    return {
      verb: 'save',
      name,
      source: { src: 'search', text: stripWrappingQuotes(tail.replace(/^search\s+/i, '').trim()) },
    };
  }
  if (tail.startsWith('"')) {
    const quoted = scanQuoted(tail, 0);
    if (!quoted || tail.slice(quoted.end).trim() !== '') return { error: usage };
    return { verb: 'save', name, source: { src: 'literal', text: quoted.value } };
  }
  const expr = parseEffectArg(tail);
  if (expr === undefined) return { error: usage };
  if (isExprParseError(expr)) return expr;
  return { verb: 'save', name, source: { src: 'expr', expr } };
}

/**
 * `share <name> = read <selector>` — identical source grammar to `save`, but the name resolves in
 * the cross-surface `/shared` handoff store (Graph app-folder) instead of the local `/work` store,
 * so another surface's session can `cat`/read it back by the same name later.
 */
function parseShare(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'share needs a name and source — usage: share <name> = read <selector> | search <text> | outline | "literal" | ($pipeline)';
  const split = splitFirstArg(rest);
  if (!split) return { error: usage };
  const name = normalizeWorkspaceName(split.arg);
  if (typeof name !== 'string') return name;
  let tail = split.tail.trim();
  if (tail.startsWith('=')) tail = tail.slice(1).trim();
  if (tail === '') return { error: usage };

  if (tail.toLowerCase() === 'outline') return { verb: 'share', name, source: { src: 'outline' } };
  if (/^read(?:\s|$)/i.test(tail)) {
    return {
      verb: 'share',
      name,
      source: { src: 'read', selector: tail.replace(/^read\s*/i, '').trim() },
    };
  }
  if (/^search\s+/i.test(tail)) {
    return {
      verb: 'share',
      name,
      source: { src: 'search', text: stripWrappingQuotes(tail.replace(/^search\s+/i, '').trim()) },
    };
  }
  if (tail.startsWith('"')) {
    const quoted = scanQuoted(tail, 0);
    if (!quoted || tail.slice(quoted.end).trim() !== '') return { error: usage };
    return { verb: 'share', name, source: { src: 'literal', text: quoted.value } };
  }
  const expr = parseEffectArg(tail);
  if (expr === undefined) return { error: usage };
  if (isExprParseError(expr)) return expr;
  return { verb: 'share', name, source: { src: 'expr', expr } };
}

function parseCat(rest: string): ParsedCommand | CommandParseError {
  const usage = 'cat needs a workspace artifact ref — usage: cat <name|ws:id> [head=N]';
  const { positional, props } = tokenizeArgs(rest);
  const ref = positional[0];
  if (!ref || positional.length > 1) return { error: usage };
  const parsedHead = parsePositiveIntProp(props.head, 'head');
  if ('error' in parsedHead) return parsedHead;
  return {
    verb: 'cat',
    ref,
    ...(parsedHead.value !== undefined ? { head: parsedHead.value } : {}),
  };
}

function parseGrep(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'grep needs an artifact ref and pattern — usage: grep <name|ws:id> "pattern" [context=N]';
  const split = splitFirstArg(rest);
  if (!split) return { error: usage };
  const ref = stripWrappingQuotes(split.arg);
  const tail = split.tail.trim();
  if (!ref || !tail) return { error: usage };

  let pattern: string;
  let propsText = '';
  if (tail.startsWith('"')) {
    const quoted = scanQuoted(tail, 0);
    if (!quoted) return { error: usage };
    pattern = quoted.value;
    propsText = tail.slice(quoted.end).trim();
  } else {
    const pieces = tail.split(/\s+/);
    pattern = pieces.shift() ?? '';
    propsText = pieces.join(' ');
  }
  if (!pattern) return { error: 'grep pattern cannot be empty' };
  const { positional, props } = tokenizeArgs(propsText);
  if (positional.length > 0) return { error: usage };
  const parsedContext = parseNonNegativeIntProp(props.context, 'context');
  if ('error' in parsedContext) return parsedContext;
  return {
    verb: 'grep',
    ref,
    pattern,
    ...(parsedContext.value !== undefined ? { context: parsedContext.value } : {}),
  };
}

/**
 * `cp <src> <dst>` — duplicate a workspace artifact under a new name (a new id; `/work` only, never
 * touches Office content). `src` is a bare ref (name or `ws:id`, resolved by `WorkspaceStore.get()`
 * at execution); `dst` is validated as a fresh artifact name with `save`'s own rules (no path
 * traversal, same character set) since it is the name a new alias is created under.
 */
function parseCp(rest: string): ParsedCommand | CommandParseError {
  const usage = 'cp needs a source and destination — usage: cp <src> <dst>';
  const { positional, props } = tokenizeArgs(rest);
  if (positional.length !== 2 || Object.keys(props).length > 0) return { error: usage };
  const dst = normalizeWorkspaceName(positional[1]!);
  if (typeof dst !== 'string') return dst;
  return { verb: 'cp', src: positional[0]!, dst };
}

/**
 * `mv <src> <dst>` — rename a workspace artifact in place (same id; `/work` only, never touches
 * Office content). Same src/dst shape as {@link parseCp}.
 */
function parseMv(rest: string): ParsedCommand | CommandParseError {
  const usage = 'mv needs a source and destination — usage: mv <src> <dst>';
  const { positional, props } = tokenizeArgs(rest);
  if (positional.length !== 2 || Object.keys(props).length > 0) return { error: usage };
  const dst = normalizeWorkspaceName(positional[1]!);
  if (typeof dst !== 'string') return dst;
  return { verb: 'mv', src: positional[0]!, dst };
}

/** `rm <name|ws:id>` — delete a workspace artifact (`/work` only, never touches Office content). */
function parseRm(rest: string): ParsedCommand | CommandParseError {
  const usage = 'rm needs an artifact ref — usage: rm <name|ws:id>';
  const { positional, props } = tokenizeArgs(rest);
  if (positional.length !== 1 || Object.keys(props).length > 0) return { error: usage };
  return { verb: 'rm', name: positional[0]! };
}

function normalizeWorkspaceName(name: string): string | CommandParseError {
  const unquoted = stripWrappingQuotes(name).trim();
  if (unquoted === '') return { error: 'workspace artifact name cannot be empty' };
  if (!/^[A-Za-z0-9][A-Za-z0-9._/-]{0,95}$/.test(unquoted)) {
    return {
      error:
        'workspace artifact name must start with a letter/number and contain only letters, numbers, ., _, -, or /',
    };
  }
  if (unquoted.includes('..') || unquoted.startsWith('/') || unquoted.endsWith('/')) {
    return { error: 'workspace artifact name must be relative and cannot contain ".."' };
  }
  return unquoted;
}

function parsePositiveIntProp(
  value: string | undefined,
  name: string,
): { value?: number } | CommandParseError {
  if (value === undefined) return {};
  const n = Number(value);
  if (!Number.isInteger(n) || n <= 0) return { error: `${name} must be a positive integer` };
  return { value: n };
}

function parseNonNegativeIntProp(
  value: string | undefined,
  name: string,
): { value?: number } | CommandParseError {
  if (value === undefined) return {};
  const n = Number(value);
  if (!Number.isInteger(n) || n < 0) return { error: `${name} must be a non-negative integer` };
  return { value: n };
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

  // Separator whitespace, then the comment body: a quoted literal OR (ADR-0005 Phase 2) an
  // effect-arg expression (`$var` / `( <pipeline> )`) evaluated against the env at dry-run.
  while (i < rest.length && /\s/.test(rest[i]!)) i++;
  if (selector === '') return { error: usage };

  const bodyRest = rest.slice(i);
  if (rest[i] !== '"') {
    // Not a quoted string → try an effect-arg expression; if it is neither, it is malformed.
    const expr = parseEffectArg(bodyRest);
    if (expr === undefined) return { error: usage };
    if (isExprParseError(expr)) return expr;
    return { verb: 'comment', selector, text: '', textExpr: expr };
  }

  const body = scanQuoted(rest, i);
  if (!body) return { error: usage };
  // Anything after the closing quote (other than whitespace) is malformed.
  if (rest.slice(body.end).trim() !== '') return { error: usage };
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
  const { positional, props } = tokenizeArgs(rest);
  if (positional.length === 0) return { error: usage };

  const range = positional[0]!;
  if (positional.length > 1) {
    const bad = positional[1]!;
    return {
      error: `format expects key=value pairs — got "${bad}" (usage: format <range> k=v)`,
    };
  }
  if (Object.keys(props).length === 0) return { error: usage };

  return { verb: 'format', range, props };
}

/**
 * Split the first command argument while respecting Excel's single-quoted sheet prefix:
 * `'Daily schedule'!B2` remains one selector. A fully quoted selector (`"Sales Q1"!A1`) is accepted
 * too. Used by verbs whose remainder is free text, so generic tokenization would lose spacing.
 */
function splitFirstArg(rest: string): { arg: string; tail: string } | undefined {
  const s = rest.trimStart();
  if (!s) return undefined;
  const m = /^"([^"]*)"(\S*)|'([^']*)'(\S*)|(\S+)/.exec(s);
  if (!m) return undefined;
  const arg =
    m[1] !== undefined
      ? `${m[1]}${m[2] ?? ''}`
      : m[3] !== undefined
        ? `'${m[3]}'${m[4] ?? ''}`
        : m[5]!;
  return { arg, tail: s.slice(m[0]!.length) };
}

/** `set A1 "hello world"` is a scalar cell value, not a request to include quote characters. */
function unquoteSetValue(value: string): string {
  return stripWrappingQuotes(value);
}

/**
 * Split a verb's argument string into positional tokens + `key=value` props, keeping `"quoted
 * values"` (with spaces) intact — used by the ADR-0007 `table`/`chart`/`cf` verbs whose props (a
 * chart `title="Top regions"`) may carry spaces. A `key="quoted"` / `key=bare` arm yields a prop;
 * anything else is positional (the range, the chart type, a bare CF mode like `databar`).
 */
function tokenizeArgs(rest: string): { positional: string[]; props: Record<string, string> } {
  const positional: string[] = [];
  const props: Record<string, string> = {};
  const re = /(\w[\w-]*)="([^"]*)"|(\w[\w-]*)=(\S+)|"([^"]*)"|'([^']*)'(\S*)|(\S+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(rest)) !== null) {
    if (m[1] !== undefined) props[m[1]] = m[2]!;
    else if (m[3] !== undefined) props[m[3]] = m[4]!;
    else if (m[5] !== undefined) positional.push(m[5]);
    else if (m[6] !== undefined) positional.push(`'${m[6]}'${m[7] ?? ''}`);
    else positional.push(m[8]!);
  }
  return { positional, props };
}

/** ADR-0007 `table <range> [headers] [name=...]` — promote a range to a native Table. */
function parseTable(rest: string): ParsedCommand | CommandParseError {
  const { positional, props } = tokenizeArgs(rest);
  const range = positional[0];
  if (range === undefined) {
    return { error: 'table needs a range — usage: table <range> [headers] [name=...]' };
  }
  // A bare `headers` flag is sugar for headers=true (the common case).
  if (positional.slice(1).includes('headers')) props.headers = 'true';
  return { verb: 'table', range, props };
}

/** ADR-0007 `chart <type> <range> [title="…"] [series=rows|columns]` — a chart over a source range. */
function parseChart(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'chart needs a type and a range — usage: chart <column|bar|line|pie|scatter|area> <range> [title="…"] [series=rows|columns]';
  const { positional, props } = tokenizeArgs(rest);
  const chartType = positional[0];
  const range = positional[1];
  if (chartType === undefined || range === undefined) return { error: usage };
  return { verb: 'chart', chartType: chartType.toLowerCase(), range, props };
}

/**
 * ADR-0007 `cf <range> <rule>` — one conditional-format rule. Tolerant of an INLINE operator
 * (`cf E2:E20 >1000 fill=#C6EFCE`), an explicit `op=/value=` form, a bare mode (`cf E2:E20 databar`),
 * or `top=N`. The parser only collects props; `compileCommand` interprets them into a typed rule.
 */
function parseCf(rest: string): ParsedCommand | CommandParseError {
  const { positional, props } = tokenizeArgs(rest);
  const range = positional[0];
  if (range === undefined) {
    return {
      error:
        'cf needs a range and a rule — usage: cf <range> >VALUE [fill=#hex] | cf <range> databar|colorscale | cf <range> top=N',
    };
  }
  for (const tok of positional.slice(1)) {
    const opMatch = /^(>=|<=|!=|>|<|=)(.+)$/.exec(tok);
    if (opMatch) {
      props.op = opMatch[1]!;
      props.value = opMatch[2]!;
    } else {
      props[tok.toLowerCase()] = 'true'; // a bare mode: databar / colorscale
    }
  }
  if (Object.keys(props).length === 0) {
    return {
      error:
        'cf needs a rule — usage: cf <range> >VALUE [fill=#hex] | cf <range> databar|colorscale | cf <range> top=N',
    };
  }
  return { verb: 'cf', range, props };
}

/**
 * `shape <pp:shape:slideId:shapeId> "text"` — replace text on one explicitly addressed
 * PowerPoint shape. The selector is intentionally host-ref shaped; a bare shape id is not enough
 * because PowerPoint shape ids are only meaningful inside a slide.
 */
function parseShape(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'shape needs a shape ref and quoted text — usage: shape <pp:shape:slideId:shapeId> "new text"';
  const quoteOffset = rest.indexOf('"');
  if (quoteOffset < 0) return { error: usage };
  const selector = rest.slice(0, quoteOffset).trim();
  if (!selector) return { error: usage };
  const quoted = scanQuoted(rest, quoteOffset);
  if (!quoted) return { error: usage };
  const text = quoted.value;
  const after = rest.slice(quoted.end).trim();
  if (after) return { error: usage };
  return { verb: 'shape', selector, text };
}

/**
 * ADR-0007 §3 `spill <range> = <expr>` — write a composed TABLE as a grid. The range is the first
 * token; the remainder (after the assignment `=`) MUST be an effect-arg expression (a `$var` or a
 * `( <pipeline> )` resolving to a table). A literal is rejected — spill is the composition sink, not
 * a verbatim writer (use `set` for a literal cell). The runtime resolves `valueExpr`→grid at dry-run.
 */
function parseSpill(rest: string): ParsedCommand | CommandParseError {
  const usage = 'spill needs a range and a table expression — usage: spill <range> = ($rows)';
  const sp = rest.search(/\s/);
  if (sp === -1) return { error: usage };
  const range = rest.slice(0, sp);
  const value = rest.slice(sp + 1).trim();
  if (range === '' || value === '') return { error: usage };
  const expr = parseEffectArg(value);
  if (expr === undefined) {
    return {
      error:
        'spill needs a composed table, not a literal — e.g. spill Report!A1 = ($rows) (use set for one cell)',
    };
  }
  if (isExprParseError(expr)) return expr;
  return { verb: 'spill', range, valueExpr: expr };
}

/** The set of valid `/`-surface capability names — the `ActuationKind` catalogue (drift-free). */
const ACTUATION_KINDS: ReadonlySet<string> = new Set(ActuationKindSchema.options);

/**
 * ADR-0008 §two-tier — `/<kind> k=v … positional …` (the specialized surface). `kind` must be an
 * `ActuationKind` (the command name IS the kind — no alias table). Args reuse the `format`-style
 * `tokenizeArgs` (quoted props + positional). An unknown kind yields a did-you-mean over the
 * catalogue. Per-surface availability (is this kind handled THIS turn) is enforced at compile, not
 * here — the parser stays structural and surface-agnostic.
 */
function parseInvoke(rawVerb: string, rest: string): ParsedCommand | CommandParseError {
  const kind = rawVerb.slice(1).toLowerCase();
  if (kind === '') {
    return { error: 'a / command needs a capability name — usage: /<capability> key=value ...' };
  }
  if (!ACTUATION_KINDS.has(kind)) {
    const near = nearestKind(kind);
    const tail = near ? ` — did you mean "/${near}"?` : '';
    return { error: `unknown capability "/${kind}"${tail} (the / surface names an ActuationKind)` };
  }
  const { positional, props } = tokenizeArgs(rest);
  return { verb: 'invoke', kind, props, args: positional };
}

/** The closest catalogue kind within an edit-distance threshold, for the `/`-surface did-you-mean. */
function nearestKind(kind: string): string | undefined {
  let best: string | undefined;
  let bestDist = Infinity;
  for (const candidate of ACTUATION_KINDS) {
    const d = levenshtein(kind, candidate);
    if (d < bestDist) {
      bestDist = d;
      best = candidate;
    }
  }
  const threshold = Math.min(4, Math.max(2, Math.ceil(kind.length / 2)));
  return best !== undefined && bestDist <= threshold ? best : undefined;
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

  // Separator whitespace, then the reply body: a quoted literal OR (ADR-0005 Phase 2) an
  // effect-arg expression (`$var` / `( <pipeline> )`) evaluated against the env at dry-run.
  while (i < rest.length && /\s/.test(rest[i]!)) i++;
  if (commentId === '') return { error: usage };

  const bodyRest = rest.slice(i);
  if (rest[i] !== '"') {
    const expr = parseEffectArg(bodyRest);
    if (expr === undefined) return { error: usage };
    if (isExprParseError(expr)) return expr;
    return { verb: 'reply', commentId, text: '', textExpr: expr };
  }

  const body = scanQuoted(rest, i);
  if (!body) return { error: usage };
  // Anything after the closing quote (other than whitespace) is malformed.
  if (rest.slice(body.end).trim() !== '') return { error: usage };
  return { verb: 'reply', commentId, text: body.value };
}

/**
 * `slide "<title>" ["<bullet>" …]` OR `slide "<title>" (<table-expr>)` (`insert-slide`). The first
 * quoted string is the title; the bullets are EITHER zero-or-more quoted strings OR a single
 * composition expression (`( <pipeline> )` / `$var`) whose resulting table's rows become bullets at
 * dry-run (`bulletsExpr`). A missing/empty title or a non-quoted, non-expression tail is corrective.
 */
function parseSlide(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'slide needs a quoted title and bullets (quoted strings or a table expression) — usage: slide "Title" "bullet one" "bullet two"  OR  slide "Title" ($rows | select a,b)';
  const t = rest.trim();
  if (!t.startsWith('"')) return { error: usage };
  const title = scanQuoted(t, 0);
  if (!title) return { error: usage };
  if (title.value === '') return { error: 'slide title cannot be empty' };

  const tail = t.slice(title.end).trim();
  if (tail === '') return { verb: 'slide', title: title.value, bullets: [] };
  // Expression bullets: a parenthesized pipeline or a bare `$var` (quoted bullets always start `"`).
  if (tail.startsWith('(') || tail.startsWith('$')) {
    const expr = parseEffectArg(tail);
    if (expr === undefined) return { error: usage };
    if (isExprParseError(expr)) return expr;
    return { verb: 'slide', title: title.value, bullets: [], bulletsExpr: expr };
  }
  const bullets = scanQuotedList(tail);
  if (!bullets) return { error: usage };
  return { verb: 'slide', title: title.value, bullets };
}

/**
 * Parse a free-text effect slot: EITHER one quoted literal (`"…"`, quote-aware, no trailing junk) OR
 * a composition expression (`( <pipeline> )` / `$var`, via {@link parseEffectArg}). Returns the
 * literal text, the parsed expression, or a corrective. Empty literals are allowed here; callers
 * that require non-empty content (mail/post) reject `''` themselves.
 */
function parseBodyArg(
  rest: string,
  usage: string,
): { text: string } | { expr: ParsedExpr } | CommandParseError {
  const t = rest.trim();
  if (t === '') return { error: usage };
  if (t.startsWith('"')) {
    const q = scanQuoted(t, 0);
    if (!q) return { error: usage };
    if (t.slice(q.end).trim() !== '') return { error: usage };
    return { text: q.value };
  }
  const expr = parseEffectArg(t);
  if (expr === undefined) return { error: usage };
  if (isExprParseError(expr)) return expr;
  return { expr };
}

/** Read a leading quoted anchor (title/subject) and return it + the remaining tail, or undefined. */
function scanLeadingQuoted(rest: string): { value: string; tail: string } | undefined {
  const t = rest.trim();
  if (!t.startsWith('"')) return undefined;
  const q = scanQuoted(t, 0);
  if (!q) return undefined;
  return { value: q.value, tail: t.slice(q.end) };
}

/**
 * `page "<title>" "<body>"` OR `page "<title>" (<expr>)` (`append-page`). A quoted page title, then
 * a body that is EITHER a quoted literal or a composition expression (evaluated to text at dry-run).
 * A missing/empty title or a malformed body is corrective.
 */
function parsePage(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'page needs a quoted title and a body (quoted or an expression) — usage: page "Title" "body text"  OR  page "Title" ($notes | ...)';
  const head = scanLeadingQuoted(rest);
  if (!head) return { error: usage };
  if (head.value === '') return { error: 'page title cannot be empty' };
  const arg = parseBodyArg(head.tail, usage);
  if ('error' in arg) return arg;
  return 'expr' in arg
    ? { verb: 'page', title: head.value, body: '', bodyExpr: arg.expr }
    : { verb: 'page', title: head.value, body: arg.text };
}

/**
 * `mail "<body>"` OR `mail (<expr>)` (`reply-mail`; `reply` is the comment-reply verb, so this is
 * `mail`). The reply body is a quoted literal or a composition expression. Missing/empty is corrective.
 */
function parseMail(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'mail needs a quoted body or an expression — usage: mail "reply body text"  OR  mail ($draft | ...)';
  const arg = parseBodyArg(rest, usage);
  if ('error' in arg) return arg;
  if ('expr' in arg) return { verb: 'mail', body: '', bodyExpr: arg.expr };
  if (arg.text === '') return { error: 'mail body cannot be empty' };
  return { verb: 'mail', body: arg.text };
}

/**
 * `post "<text>"` OR `post (<expr>)` (`post-message`). The Teams chat post text is a quoted literal
 * or a composition expression; the bridge stages it for review (never auto-sent). Missing/empty is
 * corrective.
 */
function parsePost(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'post needs a quoted text or an expression — usage: post "message text"  OR  post ($summary | ...)';
  const arg = parseBodyArg(rest, usage);
  if ('error' in arg) return arg;
  if ('expr' in arg) return { verb: 'post', text: '', textExpr: arg.expr };
  if (arg.text === '') return { error: 'post text cannot be empty' };
  return { verb: 'post', text: arg.text };
}

/**
 * `compose "<subject>" "<body>"` OR `compose "<subject>" (<expr>)` (`create-mail`). A quoted subject,
 * then a body that is a quoted literal or a composition expression. The bridge opens a fresh message
 * form (recipients left for the user); never auto-sent. A missing/empty subject is corrective.
 */
function parseCompose(rest: string): ParsedCommand | CommandParseError {
  const usage =
    'compose needs a quoted subject and a body (quoted or an expression) — usage: compose "Subject" "body text"  OR  compose "Subject" ($draft | ...)';
  const head = scanLeadingQuoted(rest);
  if (!head) return { error: usage };
  const subject = head.value;
  if (subject === '') return { error: 'compose subject cannot be empty' };
  const arg = parseBodyArg(head.tail, usage);
  if ('error' in arg) return arg;
  if ('expr' in arg) return { verb: 'compose', subject, body: '', bodyExpr: arg.expr };
  return { verb: 'compose', subject, body: arg.text };
}

/**
 * Scan a whitespace-separated list of double-quoted strings (`"a" "b" "c"`), honoring `\"`/`\\`
 * escapes via {@link scanQuoted}. Returns the unescaped values, or `null` if the input is not a
 * clean sequence of quoted strings (a bare/unquoted token, an unterminated quote, or junk between
 * strings). An empty/whitespace-only input yields `[]`.
 */
function scanQuotedList(s: string): string[] | null {
  const values: string[] = [];
  let i = 0;
  while (i < s.length) {
    while (i < s.length && /\s/.test(s[i]!)) i++;
    if (i >= s.length) break;
    if (s[i] !== '"') return null; // a bare/unquoted token between quotes
    const scanned = scanQuoted(s, i);
    if (!scanned) return null; // unterminated quote
    values.push(scanned.value);
    i = scanned.end;
  }
  return values;
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
 * (pipeline / `let`), an ADR-0005 Phase-3 skill definition (`def … end`) or call, or a corrective
 * parse error from any layer. The runtime loop dispatches on `kind` vs `verb` vs `error`.
 */
export type ProgramEntry =
  | ParsedCommand
  | ParsedExpr
  | ParsedAnalysisBinding
  | ParsedVerifiedFinish
  | ParsedSkillDef
  | ParsedSkillCall
  | CommandParseError;

/** A runtime-owned artifact binding. The name excludes `$`, like expression bindings. */
export interface ParsedAnalysisBinding {
  kind: 'analysis-binding';
  name: string;
  request: string;
}

/** Requests completion only after the runtime verifies every effect; never grants approval. */
export interface ParsedVerifiedFinish {
  kind: 'verified-finish';
}

export { ANALYSIS_BINDING_KINDS } from './analysis-actions.js';

export function isProgramAnalysisBinding(entry: ProgramEntry): entry is ParsedAnalysisBinding {
  return 'kind' in entry && entry.kind === 'analysis-binding';
}

export function isProgramVerifiedFinish(entry: ProgramEntry): entry is ParsedVerifiedFinish {
  return 'kind' in entry && entry.kind === 'verified-finish';
}

/** Recognize analysis before the expression parser, including malformed binding names. */
function parseAnalysisBinding(line: string): ParsedAnalysisBinding | CommandParseError | undefined {
  const match = /^let\s+([^=]*)=\s*analyze(?:\s+(.*))?$/i.exec(line);
  if (!match) return;
  const name = match[1]!.trim();
  if (!/^\$[A-Za-z_][A-Za-z0-9_]{0,63}$/.test(name))
    return {
      error: 'analysis binding needs a $-prefixed name — usage: let $name = analyze <JSON action>',
    };
  const command = parseCommandLine(`analyze ${match[2] ?? ''}`);
  if (isCommandParseError(command)) return command;
  if (command.verb !== 'analyze') return { error: 'analysis binding requires a JSON action' };
  const action = JSON.parse(command.request) as Record<string, unknown>;
  if (!isAnalysisBindingKind(action['kind']))
    return {
      error:
        'analysis bindings require capture, query, reconcile, filter, or inspect; effects and recovery cannot bind artifacts',
    };
  return { kind: 'analysis-binding', name: name.slice(1), request: command.request };
}

export function isProgramExpr(entry: ProgramEntry): entry is ParsedExpr {
  return 'kind' in entry && (entry.kind === 'pipeline' || entry.kind === 'let');
}

export function isProgramCommand(entry: ProgramEntry): entry is ParsedCommand {
  return 'verb' in entry;
}

export function isProgramSkillDef(entry: ProgramEntry): entry is ParsedSkillDef {
  return 'kind' in entry && entry.kind === 'skill-def';
}

export function isProgramSkillCall(entry: ProgramEntry): entry is ParsedSkillCall {
  return 'kind' in entry && entry.kind === 'skill-call';
}

/**
 * Parse a single ADR-0005 program line (already trimmed, non-blank, non-comment) into a program
 * entry, given the set of currently-registered skill names. Routing, in order:
 *   1. a registered skill name as the first token → a {@link ParsedSkillCall} (positional args);
 *   2. an expression line (top-level `|` or leading `let`) → the expression parser;
 *   3. otherwise the unchanged ADR-0004 command parser.
 *
 * A skill name is checked FIRST (before the command parser's did-you-mean) so a call never degrades
 * to an "unknown verb" error; a `def …`/`end` line is handled by the block grouper, not here.
 */
function parseProgramLine(line: string, knownSkills: ReadonlySet<string>): ProgramEntry {
  if (/^finish(?:\s|$)/i.test(line))
    return /^finish\s+when=verified$/i.test(line)
      ? { kind: 'verified-finish' }
      : { error: 'finish requires exactly when=verified — usage: finish when=verified' };
  const analysisBinding = parseAnalysisBinding(line);
  if (analysisBinding) return analysisBinding;
  const firstSpace = line.search(/\s/);
  const head = firstSpace === -1 ? line : line.slice(0, firstSpace);
  if (knownSkills.has(head)) {
    const rest = firstSpace === -1 ? '' : line.slice(firstSpace + 1).trim();
    return parseSkillCall(head, rest);
  }
  return isExpressionLine(line) ? parseExpressionLine(line) : parseCommandLine(line);
}

/**
 * Parse the whole model reply as an ADR-0005 *program*: extract the ```cmd fence, then walk the
 * lines, grouping each `def <name>(…): … end` block into ONE {@link ParsedSkillDef} entry and
 * routing every other non-blank, non-comment line through {@link parseProgramLine}.
 *
 * `knownSkills` (default empty) is the runtime's live registry of skill names; a line whose first
 * token is a registered skill becomes a {@link ParsedSkillCall}. This is a superset of
 * {@link parseCommandBlock}: with no skills and no `def`, every line parses EXACTLY as before, so
 * all ADR-0004/Phase-1/Phase-2 behavior is preserved. `found` is false when there is no fence.
 *
 * `def … end` grouping is defensive: an unterminated `def` (no `end` before the block ends) and a
 * stray `end` (no open `def`) each yield a corrective error entry — never a throw, never a silent
 * drop of the body.
 */
export function parseProgramBlock(
  modelText: string,
  knownSkills: ReadonlySet<string> = new Set(),
  options: { requireCompleteFrame?: boolean } = {},
): {
  found: boolean;
  entries: ProgramEntry[];
} {
  const block = extractCommandBlock(modelText);
  if (block === null) return { found: false, entries: [] };

  const entries: ProgramEntry[] = [];
  const lines = block.split('\n');
  let i = 0;
  let finished = false;
  while (i < lines.length) {
    const line = lines[i]!.trim();
    i++;
    if (line === '' || line.startsWith('#')) continue;
    if (finished) {
      entries.push({ error: 'finish when=verified must be the final program entry' });
      break;
    }

    if (isSkillEnd(line)) {
      // A stray `end` with no open `def` — corrective, not a silent drop.
      entries.push({ error: '`end` without a matching `def`' });
      continue;
    }
    if (isSkillDefHeader(line)) {
      const def = collectSkillDef(line, lines, i);
      entries.push(def.entry);
      i = def.nextIndex;
      continue;
    }
    const entry = parseProgramLine(line, knownSkills);
    entries.push(entry);
    finished = isProgramVerifiedFinish(entry);
  }
  // Legacy programs retain recovery from a missing fence. Verified completion requires a
  // complete response: a truncated stream must never be interpreted as a completion request.
  // Inspect the complete response, including later blocks. Otherwise a first non-finish block
  // could hide a completion request in a second block and retain its prefix effects.
  const requestsFinish =
    options.requireCompleteFrame ||
    modelText.split('\n').some((line) => /^finish(?:\s|$)/i.test(line.trim()));
  if (requestsFinish) {
    const closed = /^```cmd[^\S\n]*\r?\n[\s\S]*?^```[^\S\n]*(?:\r?\n|$)/im.test(modelText);
    const singleFrame = /^```cmd[^\S\n]*\r?\n([\s\S]*?)\r?\n```$/i.exec(modelText.trim());
    if (!closed) entries.push({ error: 'finish when=verified requires a closed cmd fence' });
    else if (!singleFrame || singleFrame[1]!.includes('```'))
      entries.push({
        error:
          'finish when=verified requires exactly one cmd fence with no surrounding text or embedded fences',
      });
  }
  // Independent writes may survive ordinary corrective errors. A program requesting verified
  // completion is atomic at the parse boundary: no prefix effect survives a malformed program.
  const errors = entries.filter((entry): entry is CommandParseError => 'error' in entry);
  if (requestsFinish && errors.length > 0) return { found: true, entries: errors };
  return { found: true, entries };
}

/**
 * Collect a `def … end` block starting at the header line (already consumed; body scanning resumes
 * at `start`). Returns the grouped {@link ParsedSkillDef} (or a corrective error entry) and the
 * index of the line AFTER the closing `end`. The body lines are kept VERBATIM (trimmed, with
 * blanks/comments skipped) for the runtime to re-parse post-substitution. A nested `def` is
 * rejected (no nested definitions this wave); an unterminated body (no `end`) is corrective.
 */
function collectSkillDef(
  header: string,
  lines: string[],
  start: number,
): { entry: ParsedSkillDef | CommandParseError; nextIndex: number } {
  const parsedHeader = parseSkillDefHeader(header);
  if ('error' in parsedHeader) {
    // Still consume the body up to `end` so the rest of the block parses cleanly after the error.
    const skipTo = scanToEnd(lines, start);
    return { entry: { error: parsedHeader.error }, nextIndex: skipTo.nextIndex };
  }

  const body: string[] = [];
  let i = start;
  while (i < lines.length) {
    const line = lines[i]!.trim();
    i++;
    if (line === '' || line.startsWith('#')) continue;
    if (isSkillEnd(line)) {
      return {
        entry: { kind: 'skill-def', name: parsedHeader.name, params: parsedHeader.params, body },
        nextIndex: i,
      };
    }
    if (isSkillDefHeader(line)) {
      const skipTo = scanToEnd(lines, i);
      return {
        entry: { error: `nested def in "${parsedHeader.name}" is not allowed` },
        nextIndex: skipTo.nextIndex,
      };
    }
    body.push(line);
  }
  // Ran off the end with no `end`.
  return {
    entry: { error: `def "${parsedHeader.name}" is missing a closing \`end\`` },
    nextIndex: i,
  };
}

/** Skip forward to (and past) the next `end` line; used to recover after a malformed `def`. */
function scanToEnd(lines: string[], start: number): { nextIndex: number } {
  let i = start;
  while (i < lines.length) {
    if (isSkillEnd(lines[i]!.trim())) return { nextIndex: i + 1 };
    i++;
  }
  return { nextIndex: i };
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
 * surface must never advertise a host read it cannot serve). `context` is the exception: it is
 * runtime-served, read-only, and always available so the model can ask the host for an upload/context
 * strategy before escalating. A WRITE verb appears ONLY when `manifest.actuations[]` contains its
 * mapped `ActuationKind`. Surface selector hints differ (Excel reads an A1/NamedRange; Word's
 * `read` is whole-document), so the smaller per-surface grammar is fewer tokens to get wrong.
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
    ls: {
      verb: 'ls',
      usage: 'ls <path>',
      hint: 'list DocFs directory entries under /doc or /work',
    },
    find: {
      verb: 'find',
      usage: 'find <path> [glob]',
      hint: 'recursively list DocFs file paths under /doc or /work, optionally filtered by a glob',
    },
    tail: {
      verb: 'tail',
      usage: 'tail <path> [n]',
      hint: 'show the last n lines (default 10) of a DocFs file under /doc or /work',
    },
    list: {
      verb: 'list',
      usage: 'list [kind]',
      hint: 'list addressable context refs without reading their content',
    },
    inspect: {
      verb: 'inspect',
      usage: 'inspect <refId|selector>',
      hint: 'materialize one context ref or selector after listing it',
    },
    properties: {
      verb: 'properties',
      usage: 'properties <refId|selector>',
      hint: 'show metadata, hostRef, anchor, and revealability without content',
    },
    comments: {
      verb: 'comments',
      usage: 'comments [refId|selector]',
      hint: 'list comment context refs when the host exposes them',
    },
    attachments: {
      verb: 'attachments',
      usage: 'attachments [refId|selector]',
      hint: 'list attachment refs when the host exposes them',
    },
    tables: {
      verb: 'tables',
      usage: 'tables [refId|selector]',
      hint: 'list table/range refs before table-specific work',
    },
    slides: {
      verb: 'slides',
      usage: 'slides [refId|selector]',
      hint: 'list slide refs before slide-specific work',
    },
    neighbors: {
      verb: 'neighbors',
      usage: 'neighbors [refId|selector]',
      hint: 'list nearby addressable context around the current selection/item',
    },
    context: {
      verb: 'context',
      usage:
        'context [incremental|inline-preferred|reference-preferred|upload-preferred|code-execution-preferred|analytical|full-scope ...]',
      hint: 'ask the host for a context/upload/code-execution strategy; read-only, never uploads by itself',
    },
    open: {
      verb: 'open',
      usage: 'open <refId|selector>',
      hint: 'navigate to an addressable host ref; navigation only, never mutates content',
    },
  };

  const specs: VerbSpec[] = [];
  for (const verb of READ_VERBS) {
    if (isRuntimeServedRead(verb, manifest) || declaredReads.has(verb))
      specs.push(readSpecByVerb[verb]);
  }

  for (const verb of WORKSPACE_VERBS)
    if (verb !== 'analyze' || manifest.surface === 'excel') specs.push(workspaceVerbSpec(verb));

  // Write verbs, gated by the advertised actuation kinds. Derived from WRITE_VERB_TO_KIND so a
  // new (deferred) write verb only needs an entry there + its kind in the manifest.
  const kinds = new Set(manifest.actuations.map((a) => a.kind));
  for (const [verb, kind] of Object.entries(WRITE_VERB_TO_KIND) as [WriteVerb, ActuationKind][]) {
    if (!kinds.has(kind)) continue;
    specs.push(writeVerbSpec(verb, isExcelLike));
  }

  // Specialized `/` surface: advertise long-tail actuation kinds that are live on this surface but
  // not already reachable via a core composable verb. The parser accepts every catalogue kind
  // structurally, but the runtime type-check still fails closed unless this turn's manifest
  // advertises the kind.
  const coreKinds = new Set<ActuationKind>(Object.values(WRITE_VERB_TO_KIND));
  for (const kind of [...kinds].sort()) {
    if (coreKinds.has(kind)) continue;
    specs.push(specializedVerbSpec(kind, manifest.surface));
  }

  specs.push(
    { verb: 'done', usage: 'done', hint: 'finish — you have completed the task' },
    { verb: 'help', usage: 'help', hint: 'list the available commands' },
  );
  return specs;
}

function isRuntimeServedRead(verb: ReadVerb, manifest: CapabilityManifest): boolean {
  if (['context', 'list', 'inspect', 'properties', 'open', 'neighbors'].includes(verb)) return true;
  const kinds = new Set(manifest.contextKinds);
  switch (verb) {
    case 'comments':
      return kinds.has('comment');
    case 'attachments':
      return kinds.has('attachment');
    case 'tables':
      return kinds.has('table') || kinds.has('range');
    case 'slides':
      return kinds.has('slide');
    default:
      return false;
  }
}

function workspaceVerbSpec(verb: WorkspaceVerb): VerbSpec {
  switch (verb) {
    case 'analyze':
      return {
        verb,
        usage: 'analyze <JSON action>',
        hint: 'Capture versioned ranges, query input artifact IDs with SELECT, reconcile exact decimals, or materialize a result through approval. Use help analyze for the schema.',
      };
    case 'workspace':
      return {
        verb,
        usage: 'workspace [name|ws:id]',
        hint: 'list local virtual artifacts or show one artifact summary; never reads or writes Office content',
      };
    case 'save':
      return {
        verb,
        usage: 'save <name> = read <selector> | search <text> | outline | "literal" | ($pipeline)',
        hint: 'store a bounded local artifact from a host read or pure computation; returns a compact handle',
      };
    case 'cat':
      return {
        verb,
        usage: 'cat <name|ws:id> [head=N]',
        hint: 'preview a bounded slice of a workspace artifact by handle',
      };
    case 'grep':
      return {
        verb,
        usage: 'grep <name|ws:id> "pattern" [context=N]',
        hint: 'search a workspace artifact locally and return compact line matches',
      };
    case 'cp':
      return {
        verb,
        usage: 'cp <src> <dst>',
        hint: 'duplicate a workspace artifact under a new name (new id); local /work only',
      };
    case 'mv':
      return {
        verb,
        usage: 'mv <src> <dst>',
        hint: 'rename a workspace artifact in place (same id); local /work only',
      };
    case 'rm':
      return {
        verb,
        usage: 'rm <name|ws:id>',
        hint: 'delete a workspace artifact; local /work only',
      };
    case 'share':
      return {
        verb,
        usage: 'share <name> = read <selector> | search <text> | outline | "literal" | ($pipeline)',
        hint: 'publish a bounded artifact to the cross-surface /shared store so another surface can read it back by name; unavailable if this session has no shared store configured',
      };
  }
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
    case 'grid':
      return {
        verb: 'grid',
        usage: 'grid <range> = "a\\tb\\nc\\td"',
        hint: 'write a rectangular TSV grid as one effect, e.g. grid Report!A1:B2 = "Region\\tRevenue\\nEast\\t100"',
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
    case 'slide':
      return {
        verb: 'slide',
        usage: 'slide "Title" "bullet" ...  OR  slide "Title" (<table expr>)',
        hint: 'add a slide; bullets can be a table expression, e.g. slide "Top accounts" ($rows | select name,arr)',
      };
    case 'page':
      return {
        verb: 'page',
        usage: 'page "Title" "body"  OR  page "Title" (<expr>)',
        hint: 'append a synthesized page, e.g. page "Meeting notes" ($decisions | head 5)',
      };
    case 'mail':
      return {
        verb: 'mail',
        usage: 'mail "body"  OR  mail (<expr>)',
        hint: 'stage a reviewable reply; body may compose, e.g. mail ($draft | head 1)',
      };
    case 'post':
      return {
        verb: 'post',
        usage: 'post "text"  OR  post (<expr>)',
        hint: 'stage a reviewable chat post, e.g. post ($summary | head 1)',
      };
    case 'compose':
      return {
        verb: 'compose',
        usage: 'compose "Subject" "body"  OR  compose "Subject" (<expr>)',
        hint: 'draft a new grounded email, e.g. compose "Follow-up on Q3" ($draft | head 1)',
      };
    case 'table':
      return {
        verb: 'table',
        usage: 'table <range> [headers] [name=...]',
        hint: 'promote a range to a native Table, e.g. table Report!A1:C12 headers',
      };
    case 'chart':
      return {
        verb: 'chart',
        usage: 'chart <column|bar|line|pie|scatter|area> <range> [title="…"] [series=rows|columns]',
        hint: 'add a chart over a range, e.g. chart column Report!A1:B11 title="Top regions"',
      };
    case 'cf':
      return {
        verb: 'cf',
        usage:
          'cf <range> >VALUE [fill=#hex]  OR  cf <range> databar|colorscale  OR  cf <range> top=N',
        hint: 'add a conditional-format rule, e.g. cf Sales!E2:E200 >100000 fill=#C6EFCE',
      };
    case 'spill':
      return {
        verb: 'spill',
        usage: 'spill <range> = (<table pipeline>)',
        hint: 'write a composed table as a grid, e.g. spill Report!A1 = ($top | select Region,Revenue)',
      };
    case 'shape':
      return {
        verb: 'shape',
        usage: 'shape <pp:shape:slideId:shapeId> "text"',
        hint: 'replace text in one selected/addressed PowerPoint shape',
      };
  }
}

function specializedVerbSpec(kind: ActuationKind, surface: Surface): VerbSpec {
  const registry = registryEntryForKindAndSurface(kind, surface);
  if (registry) {
    return {
      verb: kind,
      usage: `${registry.command} [key=value ...]`,
      hint: registry.useWhen,
    };
  }
  switch (kind) {
    case 'insert-text':
      return {
        verb: kind,
        usage: '/insert-text text="..." [match="exact anchor"] [contextHint="..."]',
        hint: 'Word: insert plain text at the selection or after an exact content anchor',
      };
    case 'replace-selection':
      return {
        verb: kind,
        usage: '/replace-selection text="..."',
        hint: 'Word: replace the current selection; fails closed when nothing is selected',
      };
    case 'insert-ooxml':
      return {
        verb: kind,
        usage: '/insert-ooxml ooxml="<w:p/>" [match="exact anchor"] [contextHint="..."]',
        hint: 'Word: insert rich OOXML at the selection or after an exact content anchor',
      };
    case 'fill-content-control':
      return {
        verb: kind,
        usage: '/fill-content-control id=<contentControlId> text="..."',
        hint: 'Word: fill a known content control id',
      };
    default:
      return {
        verb: kind,
        usage: `/${kind} [key=value ...]`,
        hint: 'invoke a specialized host capability only when this surface advertises it',
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
