import {
  EFFECT_COMPOSE_ERROR,
  EFFECT_VERBS,
  type ParsedExpr,
  type PipeSource,
} from '@ge/contracts';

/**
 * ADR-0005 (Phase 1) — the runtime value model + pure transform registry + the evaluator.
 *
 * The keystone of ADR-0005 is a typed `Value` between reads and transforms. Reads produce values
 * (a GFM table from Excel, or free text); pure transforms compose them (`filter`, `select`, `sum`,
 * `sort`, …). Phase 1 is **pure only**: a pipeline evaluates to a `Value`; `let` binds it into the
 * env. There are NO effect terminals — piping into `set`/`suggest`/`comment`/`format` is rejected
 * upstream (the loop never routes an effect verb through here). The transform registry is the
 * durable artifact: pure, total functions tested hard.
 */

/* ───────────────────────────── Value ───────────────────────────── */

export type Value =
  | { kind: 'table'; columns: string[]; rows: string[][] }
  | { kind: 'number'; value: number }
  | { kind: 'text'; value: string };

/** The corrective result a transform/evaluator returns instead of a `Value` (self-heals upstream). */
export interface EvalError {
  error: string;
}

export function isEvalError(v: Value | EvalError): v is EvalError {
  return 'error' in v;
}

/* ───────────────────────── GFM table parsing/rendering ───────────────────── */

const SEPARATOR_CELL = /^:?-{1,}:?$/;

/**
 * Parse a GFM markdown table (the shape Excel reads return — `| a | b |` rows with a
 * `| --- | --- |` separator) into a `table` Value. Tolerates leading/trailing prose around the
 * table, a missing/lenient separator row, and ragged rows (padded/truncated to the header width).
 * Returns `undefined` when no table is present (the caller falls back to a `text` Value).
 */
export function parseTable(text: string): Value | undefined {
  const rows: string[][] = [];
  for (const raw of text.split('\n')) {
    const line = raw.trim();
    // A table row is a line bounded by pipes (GFM): `| a | b |`.
    if (!(line.startsWith('|') && line.endsWith('|') && line.length > 1)) {
      // Allow blank lines / prose between the table and surrounding content only before/after;
      // once we have started a table, a non-row line ends it.
      if (rows.length > 0) break;
      continue;
    }
    rows.push(splitRow(line));
  }
  if (rows.length === 0) return undefined;

  const columns = rows[0]!;
  const body = rows.slice(1);

  // Drop a separator row (`--- | :---: | ---`) if present as the second line.
  if (body.length > 0 && body[0]!.every((c) => SEPARATOR_CELL.test(c.trim()))) {
    body.shift();
  }

  const width = columns.length;
  const normalized = body.map((r) => normalizeRow(r, width));
  return { kind: 'table', columns: columns.map((c) => c.trim()), rows: normalized };
}

/** Split a `| a | b |` line into trimmed cells (strips the leading/trailing pipe). */
function splitRow(line: string): string[] {
  const inner = line.slice(1, line.length - 1);
  return inner.split('|').map((c) => c.trim());
}

/** Pad/truncate a row to `width` cells so the table is rectangular. */
function normalizeRow(row: string[], width: number): string[] {
  const out = row.slice(0, width).map((c) => c.trim());
  while (out.length < width) out.push('');
  return out;
}

/** Max rows rendered back into a result block (keep the fed-back value token-cheap). */
const RENDER_ROW_CAP = 50;

/**
 * Render a `Value` to text for the ```result``` block fed back to the model. A table → compact
 * GFM (capped at {@link RENDER_ROW_CAP} rows, with a truncation note); number/text → plain.
 */
export function renderValue(v: Value): string {
  switch (v.kind) {
    case 'number':
      return String(v.value);
    case 'text':
      return v.value;
    case 'table': {
      const head = `| ${v.columns.join(' | ')} |`;
      const sep = `| ${v.columns.map(() => '---').join(' | ')} |`;
      const shown = v.rows.slice(0, RENDER_ROW_CAP);
      const body = shown.map((r) => `| ${r.join(' | ')} |`).join('\n');
      const lines = body === '' ? [head, sep] : [head, sep, body];
      if (v.rows.length > RENDER_ROW_CAP) {
        lines.push(`_(+${v.rows.length - RENDER_ROW_CAP} more rows)_`);
      }
      return lines.join('\n');
    }
  }
}

/* ───────────────────────────── transform registry ───────────────────────── */

/** A pure transform: an input `Value` + the verbatim arg string → a `Value` or a corrective error. */
export type Transform = (input: Value, rawArgs: string) => Value | EvalError;

/** Comparison/filter operators, longest-first so `>=`/`<=`/`!=` win over `>`/`<`/`=`. */
const FILTER_OPS = ['>=', '<=', '!=', 'contains', '=', '>', '<'] as const;
type FilterOp = (typeof FILTER_OPS)[number];

export const TRANSFORMS: Record<string, Transform> = {
  filter,
  select,
  sum: aggregate('sum'),
  avg: aggregate('avg'),
  min: aggregate('min'),
  max: aggregate('max'),
  count,
  sort,
  head: takeRows('head'),
  tail: takeRows('tail'),
  sed,
};

/** One-line usage per transform, for the COMPOSITION advertisement in the prompt. */
export const TRANSFORM_USAGE: Record<string, string> = {
  filter: 'filter <col><op><val>  (ops: = != > < >= <= contains) — keep matching rows',
  select: 'select <col,col,…> — project columns',
  sum: 'sum <col> — total a numeric column → number',
  avg: 'avg <col> — average a numeric column → number',
  min: 'min <col> — minimum of a column → number',
  max: 'max <col> — maximum of a column → number',
  count: 'count — number of rows → number',
  sort: 'sort <col> [desc] — sort rows by a column',
  head: 'head <n> — first n rows',
  tail: 'tail <n> — last n rows',
  sed: 'sed s/pattern/replacement/[g] — text/cell substitution',
};

/* ───────────────────────────── transform impls ───────────────────────────── */

function requireTable(
  input: Value,
  transform: string,
): { table: Extract<Value, { kind: 'table' }> } | EvalError {
  if (input.kind !== 'table') {
    return { error: `${transform} expects a table, got a ${input.kind}` };
  }
  return { table: input };
}

function colIndex(columns: string[], name: string): number {
  return columns.findIndex((c) => c === name);
}

/** `filter <col><op><val>` — keep rows where the predicate holds (numeric compare when both sides parse as numbers). */
function filter(input: Value, rawArgs: string): Value | EvalError {
  const got = requireTable(input, 'filter');
  if ('error' in got) return got;
  const { table } = got;

  const parsed = parseFilter(rawArgs);
  if ('error' in parsed) return parsed;
  const { col, op, val } = parsed;

  const idx = colIndex(table.columns, col);
  if (idx === -1) return { error: `filter: unknown column "${col}"` };

  const rows = table.rows.filter((r) => matches(r[idx] ?? '', op, val));
  return { kind: 'table', columns: table.columns, rows };
}

function parseFilter(rawArgs: string): { col: string; op: FilterOp; val: string } | EvalError {
  const usage = 'filter needs <col><op><val> — e.g. filter region=East or filter amount>=100';
  const text = rawArgs.trim();
  if (text === '') return { error: usage };

  for (const op of FILTER_OPS) {
    // `contains` is a word operator (space-separated); the rest are symbols.
    if (op === 'contains') {
      const m = text.match(/^(.+?)\s+contains\s+(.*)$/i);
      if (m) return { col: m[1]!.trim(), op, val: unquote(m[2]!.trim()) };
      continue;
    }
    const at = text.indexOf(op);
    if (at > 0) {
      return { col: text.slice(0, at).trim(), op, val: unquote(text.slice(at + op.length).trim()) };
    }
  }
  return { error: usage };
}

function matches(cell: string, op: FilterOp, val: string): boolean {
  const a = cell.trim();
  if (op === 'contains') return a.toLowerCase().includes(val.toLowerCase());

  const an = Number(a);
  const bn = Number(val);
  const numeric = a !== '' && val !== '' && !Number.isNaN(an) && !Number.isNaN(bn);

  switch (op) {
    case '=':
      return numeric ? an === bn : a === val;
    case '!=':
      return numeric ? an !== bn : a !== val;
    case '>':
      return numeric ? an > bn : a > val;
    case '<':
      return numeric ? an < bn : a < val;
    case '>=':
      return numeric ? an >= bn : a >= val;
    case '<=':
      return numeric ? an <= bn : a <= val;
  }
}

/** `select <col,col,…>` — project columns; error on an unknown column. */
function select(input: Value, rawArgs: string): Value | EvalError {
  const got = requireTable(input, 'select');
  if ('error' in got) return got;
  const { table } = got;

  const names = rawArgs
    .split(',')
    .map((c) => c.trim())
    .filter((c) => c !== '');
  if (names.length === 0)
    return { error: 'select needs at least one column — e.g. select region,amount' };

  const indices: number[] = [];
  for (const name of names) {
    const idx = colIndex(table.columns, name);
    if (idx === -1) return { error: `select: unknown column "${name}"` };
    indices.push(idx);
  }
  return {
    kind: 'table',
    columns: names,
    rows: table.rows.map((r) => indices.map((i) => r[i] ?? '')),
  };
}

type Agg = 'sum' | 'avg' | 'min' | 'max';

/** `sum|avg|min|max <col>` → a number Value over the column's numeric cells. */
function aggregate(kind: Agg): Transform {
  return (input, rawArgs) => {
    const got = requireTable(input, kind);
    if ('error' in got) return got;
    const { table } = got;

    const col = rawArgs.trim();
    if (col === '') return { error: `${kind} needs a column — e.g. ${kind} amount` };
    const idx = colIndex(table.columns, col);
    if (idx === -1) return { error: `${kind}: unknown column "${col}"` };

    const nums: number[] = [];
    for (const r of table.rows) {
      const n = Number((r[idx] ?? '').trim());
      if (!Number.isNaN(n) && (r[idx] ?? '').trim() !== '') nums.push(n);
    }
    if (nums.length === 0) return { error: `${kind}: column "${col}" has no numeric values` };

    let result: number;
    switch (kind) {
      case 'sum':
        result = nums.reduce((a, b) => a + b, 0);
        break;
      case 'avg':
        result = nums.reduce((a, b) => a + b, 0) / nums.length;
        break;
      case 'min':
        result = Math.min(...nums);
        break;
      case 'max':
        result = Math.max(...nums);
        break;
    }
    return { kind: 'number', value: result };
  };
}

/** `count` → a number Value (row count of a table). */
function count(input: Value, rawArgs: string): Value | EvalError {
  if (rawArgs.trim() !== '') return { error: 'count takes no arguments — just `count`' };
  const got = requireTable(input, 'count');
  if ('error' in got) return got;
  return { kind: 'number', value: got.table.rows.length };
}

/** `sort <col> [desc]` — stable sort rows by a column (numeric when the column is all-numeric). */
function sort(input: Value, rawArgs: string): Value | EvalError {
  const got = requireTable(input, 'sort');
  if ('error' in got) return got;
  const { table } = got;

  const tokens = rawArgs
    .trim()
    .split(/\s+/)
    .filter((t) => t !== '');
  if (tokens.length === 0) return { error: 'sort needs a column — e.g. sort amount desc' };
  const col = tokens[0]!;
  const desc = tokens[1]?.toLowerCase() === 'desc';
  if (tokens.length > 1 && !desc && tokens[1]?.toLowerCase() !== 'asc') {
    return { error: `sort: expected "asc" or "desc", got "${tokens[1]}"` };
  }

  const idx = colIndex(table.columns, col);
  if (idx === -1) return { error: `sort: unknown column "${col}"` };

  const allNumeric = table.rows.every((r) => {
    const c = (r[idx] ?? '').trim();
    return c === '' || !Number.isNaN(Number(c));
  });

  const sorted = table.rows
    .map((r, i) => ({ r, i }))
    .sort((x, y) => {
      const a = x.r[idx] ?? '';
      const b = y.r[idx] ?? '';
      let cmp: number;
      if (allNumeric) cmp = Number(a) - Number(b);
      else cmp = a < b ? -1 : a > b ? 1 : 0;
      if (cmp === 0) cmp = x.i - y.i; // stable
      return desc ? -cmp : cmp;
    })
    .map(({ r }) => r);

  return { kind: 'table', columns: table.columns, rows: sorted };
}

/** `head <n>` / `tail <n>` — first/last n rows of a table. */
function takeRows(which: 'head' | 'tail'): Transform {
  return (input, rawArgs) => {
    const got = requireTable(input, which);
    if ('error' in got) return got;
    const { table } = got;

    const n = Number(rawArgs.trim());
    if (rawArgs.trim() === '' || !Number.isInteger(n) || n < 0) {
      return { error: `${which} needs a non-negative integer — e.g. ${which} 5` };
    }
    const rows =
      which === 'head'
        ? table.rows.slice(0, n)
        : table.rows.slice(Math.max(0, table.rows.length - n));
    return { kind: 'table', columns: table.columns, rows };
  };
}

/** `sed 's/pattern/replacement/[g]'` — literal-or-regex substitution, table cells or text. */
function sed(input: Value, rawArgs: string): Value | EvalError {
  const m = /^s\/((?:[^\\/]|\\.)*)\/((?:[^\\/]|\\.)*)\/(g?)$/.exec(rawArgs.trim());
  if (!m) return { error: 'sed usage: sed s/pattern/replacement/[g] — e.g. sed s/Coast/Region/g' };
  const [, pattern, replacement, flags] = m;
  let re: RegExp;
  try {
    re = new RegExp(pattern!.replace(/\\\//g, '/'), flags === 'g' ? 'g' : '');
  } catch {
    return { error: `sed: invalid pattern — ${pattern}` };
  }
  const repl = replacement!.replace(/\\\//g, '/');
  if (input.kind === 'text') return { kind: 'text', value: input.value.replace(re, repl) };
  if (input.kind === 'table') {
    return {
      kind: 'table',
      columns: input.columns,
      rows: input.rows.map((row) => row.map((cell) => cell.replace(re, repl))),
    };
  }
  return { error: 'sed expects a table or text, got a number' };
}

function unquote(s: string): string {
  if (s.length >= 2) {
    const a = s[0]!;
    const b = s[s.length - 1]!;
    if ((a === '"' && b === '"') || (a === "'" && b === "'")) return s.slice(1, -1);
  }
  return s;
}

/* ───────────────────────────── the evaluator ───────────────────────────── */

/**
 * How the evaluator reaches the host for a source read. The runtime injects a closure that
 * dispatches `read`/`search`/`outline` through the existing ADR-0003 read path and returns the
 * read's TEXT (Excel reads are GFM tables; Word reads are free text). Pure transforms then
 * compose the parsed value. Injecting this makes `evalExpr` fully testable with a fake.
 */
export type RunRead = (source: Exclude<PipeSource, { src: 'var' }>) => Promise<string>;

/**
 * Evaluate a `ParsedExpr` to a `Value` (or a corrective `{ error }`), given the binding `env` and
 * a `runRead` to reach the host. Pure given `runRead`:
 *   • a `pipeline` evaluates its source (read/search/outline via `runRead` → `parseTable` or a
 *     text Value; `$var` → env lookup, error if unbound), then folds each stage through
 *     {@link TRANSFORMS} (unknown transform / wrong-kind input → corrective error).
 *   • a `let` evaluates its pipeline and binds the result into `env` (then returns the value).
 * On any stage error the fold stops and the corrective error is returned (the loop feeds it back).
 */
export async function evalExpr(
  expr: ParsedExpr,
  env: Map<string, Value>,
  runRead: RunRead,
): Promise<Value | EvalError> {
  if (expr.kind === 'let') {
    const result = await evalPipeline(expr, env, runRead);
    if (isEvalError(result)) return result;
    env.set(expr.name, result);
    return result;
  }
  return evalPipeline(expr, env, runRead);
}

async function evalPipeline(
  expr: ParsedExpr,
  env: Map<string, Value>,
  runRead: RunRead,
): Promise<Value | EvalError> {
  const pipeline = expr.kind === 'let' ? expr.pipeline : expr;

  let current = await evalSource(pipeline.source, env, runRead);
  if (isEvalError(current)) return current;

  for (const stage of pipeline.stages) {
    // A pipe INTO an effect (`read X | set …`) — Phase 1 is pure-only; corrective, never a write.
    // Reuses the single effect-verb set from contracts so it stays in sync with the source-position guard.
    if (EFFECT_VERBS.has(stage.name)) return { error: EFFECT_COMPOSE_ERROR };
    const transform = TRANSFORMS[stage.name];
    if (!transform) {
      return {
        error: `unknown transform "${stage.name}" — available: ${Object.keys(TRANSFORMS).join(', ')}`,
      };
    }
    const next = transform(current, stage.args);
    if (isEvalError(next)) return next;
    current = next;
  }
  return current;
}

async function evalSource(
  source: PipeSource,
  env: Map<string, Value>,
  runRead: RunRead,
): Promise<Value | EvalError> {
  if (source.src === 'var') {
    const bound = env.get(source.name);
    if (!bound) return { error: `unbound variable "$${source.name}"` };
    return bound;
  }
  const text = await runRead(source);
  return parseTable(text) ?? { kind: 'text', value: text };
}

/** The set of transform names runtime exposes (for advertisement + validation). */
export const TRANSFORM_NAMES = Object.keys(TRANSFORMS);
