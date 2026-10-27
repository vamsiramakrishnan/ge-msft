import {
  EFFECT_COMPOSE_ERROR,
  EFFECT_VERBS,
  scanQuoted,
  type HostPipeSource,
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

/**
 * What a transform may see besides its input: the binding env (read-only), so a relational
 * transform (`join $other …`) can reach a second bound table. Never the host.
 */
export interface TransformContext {
  env: ReadonlyMap<string, Value>;
}

/** A pure transform: an input `Value` + the verbatim arg string (+ context) → a `Value` or a corrective error. */
export type Transform = (
  input: Value,
  rawArgs: string,
  ctx?: TransformContext,
) => Value | EvalError;

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
  derive,
  join,
  group,
  distinct,
  rename,
};

/** One-line usage per transform, for the COMPOSITION advertisement in the prompt. */
export const TRANSFORM_USAGE: Record<string, string> = {
  filter:
    'filter <pred>  — pred: <col><op><val> (ops: = != > < >= <= contains), combined with and / or / not and ( ) — keep matching rows',
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
  derive:
    'derive <newCol> = <col|num> <+|-|*|/> <col|num>  or  derive <newCol> = <fn>(<args>) — fn: upper lower trim len concat round year month day datediff — append a computed column',
  join: 'join $other on <col>[=<otherCol>] [left|inner] — combine rows with a bound table (inner by default)',
  group:
    'group <col,…> agg <fn>(<col>) as <name>[, …] — fn: sum avg min max count — one row per group',
  distinct: 'distinct [<col,…>] — unique rows (or unique values of the listed columns)',
  rename: 'rename <old>=<new>[, …] — rename columns',
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

/**
 * `filter <pred>` — keep rows where the predicate holds (numeric compare when both sides parse as
 * numbers). A predicate is one `<col><op><val>` comparison, or several combined with `and` / `or`
 * / `not` and parentheses (`not` binds tightest, then `and`, then `or`). A value containing one of
 * those words must be quoted (`filter title contains "rock and roll"`).
 */
function filter(input: Value, rawArgs: string): Value | EvalError {
  const got = requireTable(input, 'filter');
  if ('error' in got) return got;
  const { table } = got;

  const pred = parsePredicate(rawArgs);
  if ('error' in pred) return pred;
  const bound = bindPredicate(pred, table.columns);
  if ('error' in bound) return bound;

  const rows = table.rows.filter((r) => evalPredicate(bound, r));
  return { kind: 'table', columns: table.columns, rows };
}

/* ───────────────────────── boolean predicates (C1) ───────────────────────── */

type Predicate =
  | { kind: 'cmp'; col: string; op: FilterOp; val: string }
  | { kind: 'and' | 'or'; left: Predicate; right: Predicate }
  | { kind: 'not'; inner: Predicate };

type BoundPredicate =
  | { kind: 'cmp'; idx: number; op: FilterOp; val: string }
  | { kind: 'and' | 'or'; left: BoundPredicate; right: BoundPredicate }
  | { kind: 'not'; inner: BoundPredicate };

type PredToken =
  | { t: 'lp' }
  | { t: 'rp' }
  | { t: 'and' }
  | { t: 'or' }
  | { t: 'not' }
  | { t: 'atom'; text: string };

/** Tokenize a predicate: parens, the keywords and/or/not, and atoms (quoted spans kept whole). */
function tokenizePredicate(raw: string): PredToken[] | EvalError {
  const out: PredToken[] = [];
  let i = 0;
  while (i < raw.length) {
    const ch = raw[i]!;
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === '(') {
      out.push({ t: 'lp' });
      i++;
      continue;
    }
    if (ch === ')') {
      out.push({ t: 'rp' });
      i++;
      continue;
    }
    if (ch === '"') {
      const q = scanQuoted(raw, i);
      if (!q) return { error: 'filter: unterminated quote in predicate' };
      out.push({ t: 'atom', text: raw.slice(i, q.end) });
      i = q.end;
      continue;
    }
    let j = i;
    while (j < raw.length && !/[\s()"]/.test(raw[j]!)) j++;
    const word = raw.slice(i, j);
    const lower = word.toLowerCase();
    if (lower === 'and' || lower === 'or' || lower === 'not') out.push({ t: lower });
    else out.push({ t: 'atom', text: word });
    i = j;
  }
  return out;
}

/** Recursive descent: expr := term (or term)* ; term := factor (and factor)* ; factor := not factor | ( expr ) | atoms. */
function parsePredicate(raw: string): Predicate | EvalError {
  const usage =
    'filter needs <col><op><val> — e.g. filter region=East, filter amount>=100 and region!=West, filter not (region=East or region=West)';
  const tokens = tokenizePredicate(raw);
  if ('error' in tokens) return tokens;
  if (tokens.length === 0) return { error: usage };
  let pos = 0;
  const peek = (): PredToken | undefined => tokens[pos];

  const parseExpr = (): Predicate | EvalError => {
    let left = parseTerm();
    if ('error' in left) return left;
    while (peek()?.t === 'or') {
      pos++;
      const right = parseTerm();
      if ('error' in right) return right;
      left = { kind: 'or', left, right };
    }
    return left;
  };
  const parseTerm = (): Predicate | EvalError => {
    let left = parseFactor();
    if ('error' in left) return left;
    while (peek()?.t === 'and') {
      pos++;
      const right = parseFactor();
      if ('error' in right) return right;
      left = { kind: 'and', left, right };
    }
    return left;
  };
  const parseFactor = (): Predicate | EvalError => {
    const tok = peek();
    if (!tok) return { error: `filter: predicate ends early — ${usage}` };
    if (tok.t === 'not') {
      pos++;
      const inner = parseFactor();
      if ('error' in inner) return inner;
      return { kind: 'not', inner };
    }
    if (tok.t === 'lp') {
      pos++;
      const inner = parseExpr();
      if ('error' in inner) return inner;
      if (peek()?.t !== 'rp') return { error: 'filter: missing ) in predicate' };
      pos++;
      return inner;
    }
    if (tok.t !== 'atom')
      return {
        error: `filter: unexpected "${tok.t === 'rp' ? ')' : tok.t}" in predicate — ${usage}`,
      };
    const atoms: string[] = [];
    while (peek()?.t === 'atom') {
      atoms.push((peek() as { t: 'atom'; text: string }).text);
      pos++;
    }
    const simple = parseFilter(atoms.join(' '));
    if ('error' in simple) return simple;
    return { kind: 'cmp', ...simple };
  };

  const pred = parseExpr();
  if ('error' in pred) return pred;
  if (pos < tokens.length) {
    const extra = tokens[pos]!;
    return {
      error: `filter: unexpected "${extra.t === 'atom' ? extra.text : extra.t === 'rp' ? ')' : extra.t}" after predicate — ${usage}`,
    };
  }
  return pred;
}

/** Resolve every column name in a predicate to an index (unknown column → corrective error). */
function bindPredicate(pred: Predicate, columns: string[]): BoundPredicate | EvalError {
  switch (pred.kind) {
    case 'cmp': {
      const idx = colIndex(columns, pred.col);
      if (idx === -1) return { error: `filter: unknown column "${pred.col}"` };
      return { kind: 'cmp', idx, op: pred.op, val: pred.val };
    }
    case 'not': {
      const inner = bindPredicate(pred.inner, columns);
      if ('error' in inner) return inner;
      return { kind: 'not', inner };
    }
    default: {
      const left = bindPredicate(pred.left, columns);
      if ('error' in left) return left;
      const right = bindPredicate(pred.right, columns);
      if ('error' in right) return right;
      return { kind: pred.kind, left, right };
    }
  }
}

function evalPredicate(pred: BoundPredicate, row: string[]): boolean {
  switch (pred.kind) {
    case 'cmp':
      return matches(row[pred.idx] ?? '', pred.op, pred.val);
    case 'not':
      return !evalPredicate(pred.inner, row);
    case 'and':
      return evalPredicate(pred.left, row) && evalPredicate(pred.right, row);
    case 'or':
      return evalPredicate(pred.left, row) || evalPredicate(pred.right, row);
  }
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

/** `<newCol> = <col|num> <op> <col|num>` where `<op>` is one of `+ - * /`. */
const DERIVE_RE = /^(\S+)\s*=\s*(\S+)\s*([+\-*/])\s*(\S+)$/;

/**
 * `derive <newCol> = <col|num> <+|-|*|/> <col|num>` — append a new column computed by one
 * arithmetic operation over two existing numeric columns (or a column and a literal), per row.
 * Deliberately minimal (no chained expressions, no functions): the one concrete `awk`-parity gap,
 * not a general expression language. Numeric coercion matches `aggregate`'s convention — a
 * trimmed empty cell is NOT treated as `0` (unlike bare `Number('')`), so a blank cell is a
 * corrective error rather than a silently-wrong zero.
 */
function derive(input: Value, rawArgs: string): Value | EvalError {
  const got = requireTable(input, 'derive');
  if ('error' in got) return got;
  const { table } = got;

  const fn = DERIVE_FN_RE.exec(rawArgs.trim());
  if (fn) return deriveFunction(table, fn[1]!, fn[2]!.toLowerCase(), fn[3]!);

  const m = DERIVE_RE.exec(rawArgs.trim());
  if (!m) return { error: DERIVE_USAGE };
  const [, newCol, lhsRaw, op, rhsRaw] = m;
  const { columns, rows } = table;

  const resolve = (raw: string, rowIdx: number): number | EvalError => {
    const idx = colIndex(columns, raw);
    if (idx === -1) {
      const n = Number(raw);
      if (Number.isNaN(n)) return { error: `derive: unknown column or number "${raw}"` };
      return n;
    }
    const cell = (rows[rowIdx]![idx] ?? '').trim();
    const n = Number(cell);
    if (cell === '' || Number.isNaN(n)) {
      return { error: `derive: non-numeric cell in column "${raw}" (row ${rowIdx + 1})` };
    }
    return n;
  };

  const outRows: string[][] = [];
  for (let i = 0; i < rows.length; i++) {
    const lhs = resolve(lhsRaw!, i);
    if (typeof lhs !== 'number') return lhs;
    const rhs = resolve(rhsRaw!, i);
    if (typeof rhs !== 'number') return rhs;
    // NOTE: division by zero is NOT special-cased — `lhs / 0` yields `Infinity`/`-Infinity` (or
    // `NaN` for `0 / 0`), and `String(...)` renders those as the literal cells "Infinity"/"NaN".
    // Value's table cells are untyped strings, so this doesn't crash or corrupt the table, but a
    // downstream `sum`/`avg` over such a column would silently propagate `Infinity`, and `NaN`
    // cells would be silently dropped (aggregate's `Number.isNaN` guard excludes them). Flagged
    // here rather than fixed: rejecting it would require a policy call (error the whole derive?
    // just that row? coerce to a sentinel?) that's out of scope for this minimal transform.
    const value =
      op === '+' ? lhs + rhs : op === '-' ? lhs - rhs : op === '*' ? lhs * rhs : lhs / rhs;
    outRows.push([...rows[i]!, String(value)]);
  }
  return { kind: 'table', columns: [...columns, newCol!], rows: outRows };
}

/* ───────────────────────── derive functions (C1) ───────────────────────── */

const DERIVE_USAGE =
  'derive usage: derive <newCol> = <col|num> <+|-|*|/> <col|num>  or  derive <newCol> = <fn>(<args>) with fn ∈ upper lower trim len concat round year month day datediff';

/** `<newCol> = <fn>(<args>)` — one function call over columns and literals. */
const DERIVE_FN_RE = /^(\S+)\s*=\s*([A-Za-z]+)\((.*)\)$/;

const DERIVE_FUNCTIONS = new Set([
  'upper',
  'lower',
  'trim',
  'len',
  'concat',
  'round',
  'year',
  'month',
  'day',
  'datediff',
]);

/** A function argument: a column (resolved per row), a quoted text literal, or a number literal. */
type DeriveArg = { kind: 'col'; idx: number; name: string } | { kind: 'lit'; text: string };

/** Split `a, "x, y", 2` on commas outside quotes. */
function splitArgs(raw: string): string[] {
  const out: string[] = [];
  let buf = '';
  let i = 0;
  while (i < raw.length) {
    const ch = raw[i]!;
    if (ch === '"') {
      const q = scanQuoted(raw, i);
      if (q) {
        buf += raw.slice(i, q.end);
        i = q.end;
        continue;
      }
    }
    if (ch === ',') {
      out.push(buf.trim());
      buf = '';
    } else buf += ch;
    i++;
  }
  if (buf.trim() !== '' || out.length > 0) out.push(buf.trim());
  return out.filter((a) => a !== '');
}

function parseDeriveArgs(raw: string, columns: string[]): DeriveArg[] | EvalError {
  const args: DeriveArg[] = [];
  for (const a of splitArgs(raw)) {
    if (a.startsWith('"')) {
      const q = scanQuoted(a, 0);
      if (!q || q.end !== a.length) return { error: `derive: malformed text literal ${a}` };
      args.push({ kind: 'lit', text: q.value });
      continue;
    }
    const idx = colIndex(columns, a);
    if (idx !== -1) {
      args.push({ kind: 'col', idx, name: a });
      continue;
    }
    if (/^-?\d+(?:\.\d+)?$/.test(a)) {
      args.push({ kind: 'lit', text: a });
      continue;
    }
    return { error: `derive: unknown column or literal "${a}" (quote text literals)` };
  }
  return args;
}

/** Parse a cell as a date; an ISO date-only cell is read as local midnight so day math is stable. */
function parseCellDate(cell: string): Date | undefined {
  const t = cell.trim();
  if (t === '') return undefined;
  const d = /^\d{4}-\d{2}-\d{2}$/.test(t) ? new Date(`${t}T00:00:00`) : new Date(t);
  return Number.isNaN(d.getTime()) ? undefined : d;
}

function deriveFunction(
  table: Extract<Value, { kind: 'table' }>,
  newCol: string,
  fn: string,
  rawArgs: string,
): Value | EvalError {
  if (!DERIVE_FUNCTIONS.has(fn))
    return { error: `derive: unknown function "${fn}" — ${DERIVE_USAGE}` };
  const args = parseDeriveArgs(rawArgs, table.columns);
  if ('error' in args) return args;

  const arity: Record<string, [number, number]> = {
    upper: [1, 1],
    lower: [1, 1],
    trim: [1, 1],
    len: [1, 1],
    concat: [1, Infinity],
    round: [1, 2],
    year: [1, 1],
    month: [1, 1],
    day: [1, 1],
    datediff: [2, 2],
  };
  const [lo, hi] = arity[fn]!;
  if (args.length < lo || args.length > hi) {
    return {
      error: `derive: ${fn} takes ${lo === hi ? lo : `${lo}–${hi === Infinity ? 'n' : hi}`} argument(s)`,
    };
  }

  const outRows: string[][] = [];
  for (let i = 0; i < table.rows.length; i++) {
    const row = table.rows[i]!;
    const text = (a: DeriveArg): string => (a.kind === 'lit' ? a.text : (row[a.idx] ?? ''));
    const num = (a: DeriveArg): number | EvalError => {
      const t = text(a).trim();
      const n = Number(t);
      if (t === '' || Number.isNaN(n)) {
        return {
          error: `derive: non-numeric value for ${fn} in column "${a.kind === 'col' ? a.name : a.text}" (row ${i + 1})`,
        };
      }
      return n;
    };
    const date = (a: DeriveArg): Date | EvalError => {
      const d = parseCellDate(text(a));
      if (!d) return { error: `derive: "${text(a)}" is not a date for ${fn} (row ${i + 1})` };
      return d;
    };
    let value: string;
    switch (fn) {
      case 'upper':
        value = text(args[0]!).toUpperCase();
        break;
      case 'lower':
        value = text(args[0]!).toLowerCase();
        break;
      case 'trim':
        value = text(args[0]!).trim();
        break;
      case 'len':
        value = String(text(args[0]!).length);
        break;
      case 'concat':
        value = args.map(text).join('');
        break;
      case 'round': {
        const n = num(args[0]!);
        if (typeof n !== 'number') return n;
        const digits = args[1] ? num(args[1]) : 0;
        if (typeof digits !== 'number') return digits;
        if (!Number.isInteger(digits) || digits < 0 || digits > 15)
          return { error: 'derive: round digits must be an integer 0–15' };
        const scale = 10 ** digits;
        value = String(Math.round(n * scale) / scale);
        break;
      }
      case 'year':
      case 'month':
      case 'day': {
        const d = date(args[0]!);
        if (!(d instanceof Date)) return d;
        value = String(
          fn === 'year' ? d.getFullYear() : fn === 'month' ? d.getMonth() + 1 : d.getDate(),
        );
        break;
      }
      case 'datediff': {
        const a = date(args[0]!);
        if (!(a instanceof Date)) return a;
        const b = date(args[1]!);
        if (!(b instanceof Date)) return b;
        value = String(Math.round((a.getTime() - b.getTime()) / 86_400_000));
        break;
      }
      default:
        return { error: `derive: unknown function "${fn}"` };
    }
    outRows.push([...row, value]);
  }
  return { kind: 'table', columns: [...table.columns, newCol], rows: outRows };
}

/* ───────────────────────── relational transforms (C1) ───────────────────────── */

const JOIN_RE = /^\$([A-Za-z_][A-Za-z0-9_]*)\s+on\s+([^\s=]+)(?:=([^\s=]+))?(?:\s+(left|inner))?$/i;

/** Numeric-aware key equality, matching `filter`'s `=` convention. */
function keysEqual(a: string, b: string): boolean {
  const x = a.trim();
  const y = b.trim();
  const xn = Number(x);
  const yn = Number(y);
  if (x !== '' && y !== '' && !Number.isNaN(xn) && !Number.isNaN(yn)) return xn === yn;
  return x === y;
}

/**
 * `join $other on <col>[=<otherCol>] [left|inner]` — combine each input row with the rows of a
 * BOUND table whose key matches. Inner (default) drops unmatched input rows; `left` keeps them with
 * blank right-hand cells. Output columns: the input's, then the other table's minus its key; a
 * right-hand name that collides with an input column gets a `_2` suffix.
 */
function join(input: Value, rawArgs: string, ctx?: TransformContext): Value | EvalError {
  const got = requireTable(input, 'join');
  if ('error' in got) return got;
  const { table } = got;
  const m = JOIN_RE.exec(rawArgs.trim());
  if (!m) return { error: 'join usage: join $other on <col>[=<otherCol>] [left|inner]' };
  const [, otherName, col, otherColRaw, modeRaw] = m;
  const other = ctx?.env.get(otherName!);
  if (!other) return { error: `join: unbound variable "$${otherName}"` };
  if (other.kind !== 'table')
    return { error: `join: $${otherName} is a ${other.kind}, not a table` };
  const otherCol = otherColRaw ?? col!;
  const left = (modeRaw ?? 'inner').toLowerCase() === 'left';

  const li = colIndex(table.columns, col!);
  if (li === -1) return { error: `join: unknown column "${col}" in the input table` };
  const ri = colIndex(other.columns, otherCol);
  if (ri === -1) return { error: `join: unknown column "${otherCol}" in $${otherName}` };

  const rightCols = other.columns.filter((_, i) => i !== ri);
  const columns = [
    ...table.columns,
    ...rightCols.map((c) => (table.columns.includes(c) ? `${c}_2` : c)),
  ];
  const blank = rightCols.map(() => '');
  const rows: string[][] = [];
  for (const row of table.rows) {
    const key = row[li] ?? '';
    let matched = false;
    for (const orow of other.rows) {
      if (!keysEqual(key, orow[ri] ?? '')) continue;
      matched = true;
      rows.push([...row, ...orow.filter((_, i) => i !== ri)]);
    }
    if (!matched && left) rows.push([...row, ...blank]);
  }
  return { kind: 'table', columns, rows };
}

const GROUP_AGG_RE = /^(sum|avg|min|max|count)\(\s*([^)]*?)\s*\)\s+as\s+(\S+)$/i;

/**
 * `group <col,…> agg <fn>(<col>) as <name>[, …]` — one output row per distinct combination of the
 * group columns (first-appearance order), followed by each aggregate. `count(*)` counts rows;
 * `count(col)` counts non-empty cells; `sum/avg/min/max` skip non-numeric cells and fail when a
 * group has none.
 */
function group(input: Value, rawArgs: string): Value | EvalError {
  const got = requireTable(input, 'group');
  if ('error' in got) return got;
  const { table } = got;
  const usage =
    'group usage: group <col,…> agg <fn>(<col>) as <name>[, …] — fn: sum avg min max count';
  const split = /^(.*?)\s+agg\s+(.+)$/i.exec(rawArgs.trim());
  if (!split) return { error: usage };
  const byNames = split[1]!
    .split(',')
    .map((c) => c.trim())
    .filter((c) => c !== '');
  if (byNames.length === 0) return { error: usage };
  const byIdx: number[] = [];
  for (const name of byNames) {
    const idx = colIndex(table.columns, name);
    if (idx === -1) return { error: `group: unknown column "${name}"` };
    byIdx.push(idx);
  }
  const aggs: Array<{ fn: string; idx: number; name: string }> = [];
  for (const spec of split[2]!
    .split(',')
    .map((a) => a.trim())
    .filter((a) => a !== '')) {
    const m = GROUP_AGG_RE.exec(spec);
    if (!m) return { error: `group: malformed aggregate "${spec}" — ${usage}` };
    const fn = m[1]!.toLowerCase();
    const col = m[2]!;
    let idx = -1;
    if (col === '*') {
      if (fn !== 'count') return { error: `group: ${fn}(*) is not allowed — name a column` };
    } else {
      idx = colIndex(table.columns, col);
      if (idx === -1) return { error: `group: unknown column "${col}"` };
    }
    const name = m[3]!;
    if (byNames.includes(name) || aggs.some((a) => a.name === name)) {
      return { error: `group: duplicate output column "${name}"` };
    }
    aggs.push({ fn, idx, name });
  }
  if (aggs.length === 0) return { error: usage };

  const buckets = new Map<string, { key: string[]; rows: string[][] }>();
  for (const row of table.rows) {
    const key = byIdx.map((i) => row[i] ?? '');
    const k = key.join('\u0000');
    const bucket = buckets.get(k);
    if (bucket) bucket.rows.push(row);
    else buckets.set(k, { key, rows: [row] });
  }

  const rows: string[][] = [];
  for (const { key, rows: members } of buckets.values()) {
    const out = [...key];
    for (const agg of aggs) {
      if (agg.fn === 'count') {
        out.push(
          String(
            agg.idx === -1
              ? members.length
              : members.filter((r) => (r[agg.idx] ?? '').trim() !== '').length,
          ),
        );
        continue;
      }
      const nums = members
        .map((r) => (r[agg.idx] ?? '').trim())
        .filter((c) => c !== '' && !Number.isNaN(Number(c)))
        .map(Number);
      if (nums.length === 0) {
        return {
          error: `group: ${agg.fn}(${table.columns[agg.idx]}) has no numeric values in group ${key.join('/')}`,
        };
      }
      const total = nums.reduce((a, b) => a + b, 0);
      const v =
        agg.fn === 'sum'
          ? total
          : agg.fn === 'avg'
            ? total / nums.length
            : agg.fn === 'min'
              ? Math.min(...nums)
              : Math.max(...nums);
      out.push(String(v));
    }
    rows.push(out);
  }
  return { kind: 'table', columns: [...byNames, ...aggs.map((a) => a.name)], rows };
}

/** `distinct [<col,…>]` — unique rows; with columns, the unique values of that projection. */
function distinct(input: Value, rawArgs: string): Value | EvalError {
  const got = requireTable(input, 'distinct');
  if ('error' in got) return got;
  let { table } = got;
  if (rawArgs.trim() !== '') {
    const projected = select(table, rawArgs);
    if ('error' in projected) return { error: projected.error.replace(/^select/, 'distinct') };
    table = projected as Extract<Value, { kind: 'table' }>;
  }
  const seen = new Set<string>();
  const rows: string[][] = [];
  for (const row of table.rows) {
    const k = row.join('\u0000');
    if (seen.has(k)) continue;
    seen.add(k);
    rows.push(row);
  }
  return { kind: 'table', columns: table.columns, rows };
}

/** `rename <old>=<new>[, …]` — rename columns (unknown source / duplicate target are errors). */
function rename(input: Value, rawArgs: string): Value | EvalError {
  const got = requireTable(input, 'rename');
  if ('error' in got) return got;
  const { table } = got;
  const pairs = rawArgs
    .split(',')
    .map((p) => p.trim())
    .filter((p) => p !== '');
  if (pairs.length === 0) return { error: 'rename usage: rename <old>=<new>[, <old>=<new>]' };
  const columns = [...table.columns];
  for (const pair of pairs) {
    const eq = pair.indexOf('=');
    const from = eq === -1 ? '' : pair.slice(0, eq).trim();
    const to = eq === -1 ? '' : pair.slice(eq + 1).trim();
    if (from === '' || to === '')
      return { error: `rename: malformed pair "${pair}" — use <old>=<new>` };
    const idx = colIndex(columns, from);
    if (idx === -1) return { error: `rename: unknown column "${from}"` };
    if (columns.includes(to)) return { error: `rename: column "${to}" already exists` };
    columns[idx] = to;
  }
  return { kind: 'table', columns, rows: table.rows };
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
export type RunRead = (source: HostPipeSource) => Promise<string>;

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
    const next = transform(current, stage.args, { env });
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
  // C1 literals — values the program itself carries; never a host read.
  if (source.src === 'number') return { kind: 'number', value: source.value };
  if (source.src === 'text') return { kind: 'text', value: source.value };
  if (source.src === 'table') return { kind: 'table', columns: source.columns, rows: source.rows };
  const text = await runRead(source);
  return parseTable(text) ?? { kind: 'text', value: text };
}

/** The set of transform names runtime exposes (for advertisement + validation). */
export const TRANSFORM_NAMES = Object.keys(TRANSFORMS);
