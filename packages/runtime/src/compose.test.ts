import { describe, it, expect } from 'vitest';
import { parseExpressionLine, TRANSFORM_NAMES, type PipeSource } from '@ge/contracts';
import {
  TRANSFORMS,
  parseTable,
  renderValue,
  evalExpr,
  isEvalError,
  type Value,
  type RunRead,
} from './compose.js';

describe('TRANSFORMS ↔ TRANSFORM_NAMES (no drift)', () => {
  it('the registry provides exactly the names the grammar boundary advertises', () => {
    expect(Object.keys(TRANSFORMS).sort()).toEqual([...TRANSFORM_NAMES].sort());
  });
});

const table = (columns: string[], rows: string[][]): Value => ({ kind: 'table', columns, rows });

const SALES = table(
  ['region', 'amount', 'rep'],
  [
    ['East', '100', 'Ann'],
    ['West', '250', 'Bob'],
    ['East', '50', 'Cy'],
    ['North', '300', 'Dee'],
  ],
);

/* ───────────────────────────── parseTable / renderValue ───────────────────── */

describe('parseTable', () => {
  it('round-trips a GFM table (with the --- separator row)', () => {
    const gfm = '| region | amount |\n| --- | --- |\n| East | 100 |\n| West | 250 |';
    expect(parseTable(gfm)).toEqual(
      table(
        ['region', 'amount'],
        [
          ['East', '100'],
          ['West', '250'],
        ],
      ),
    );
  });

  it('tolerates surrounding prose and a colon-aligned separator', () => {
    const gfm = 'Here is the data:\n| a | b |\n| :--- | ---: |\n| 1 | 2 |\n\ndone.';
    expect(parseTable(gfm)).toEqual(table(['a', 'b'], [['1', '2']]));
  });

  it('pads ragged rows to the header width', () => {
    const gfm = '| a | b | c |\n| --- | --- | --- |\n| 1 | 2 |';
    expect(parseTable(gfm)).toEqual(table(['a', 'b', 'c'], [['1', '2', '']]));
  });

  it('returns undefined for non-table text', () => {
    expect(parseTable('just a sentence, no table.')).toBeUndefined();
    expect(parseTable('')).toBeUndefined();
  });
});

describe('renderValue', () => {
  it('renders a number / text plainly', () => {
    expect(renderValue({ kind: 'number', value: 42 })).toBe('42');
    expect(renderValue({ kind: 'text', value: 'hello' })).toBe('hello');
  });

  it('renders a table back to compact GFM', () => {
    expect(renderValue(table(['a', 'b'], [['1', '2']]))).toBe(
      '| a | b |\n| --- | --- |\n| 1 | 2 |',
    );
  });

  it('round-trips parse → render → parse', () => {
    const v = parseTable(renderValue(SALES))!;
    expect(v).toEqual(SALES);
  });
});

/* ───────────────────────────── transforms ───────────────────────────── */

describe('TRANSFORMS — filter', () => {
  it('string equality', () => {
    const out = TRANSFORMS.filter!(SALES, 'region=East') as Extract<Value, { kind: 'table' }>;
    expect(out.rows.map((r) => r[2])).toEqual(['Ann', 'Cy']);
  });

  it('numeric comparison when both sides parse as numbers', () => {
    const out = TRANSFORMS.filter!(SALES, 'amount>=250') as Extract<Value, { kind: 'table' }>;
    expect(out.rows.map((r) => r[1])).toEqual(['250', '300']);
  });

  it('string comparison falls back when not numeric', () => {
    const out = TRANSFORMS.filter!(SALES, 'rep>Bob') as Extract<Value, { kind: 'table' }>;
    expect(out.rows.map((r) => r[2])).toEqual(['Cy', 'Dee']);
  });

  it('!= and contains', () => {
    expect(
      (TRANSFORMS.filter!(SALES, 'region!=East') as Extract<Value, { kind: 'table' }>).rows,
    ).toHaveLength(2);
    expect(
      (TRANSFORMS.filter!(SALES, 'rep contains a') as Extract<Value, { kind: 'table' }>).rows.map(
        (r) => r[2],
      ),
    ).toEqual(['Ann']);
  });

  it('errors on an unknown column or wrong input kind', () => {
    expect(TRANSFORMS.filter!(SALES, 'nope=1')).toMatchObject({ error: expect.any(String) });
    expect(TRANSFORMS.filter!({ kind: 'number', value: 1 }, 'a=1')).toMatchObject({
      error: expect.stringContaining('table'),
    });
  });
});

describe('TRANSFORMS — select', () => {
  it('projects columns in order', () => {
    const out = TRANSFORMS.select!(SALES, 'rep, region') as Extract<Value, { kind: 'table' }>;
    expect(out.columns).toEqual(['rep', 'region']);
    expect(out.rows[0]).toEqual(['Ann', 'East']);
  });

  it('errors on an unknown column', () => {
    expect(TRANSFORMS.select!(SALES, 'region,ghost')).toMatchObject({
      error: expect.stringContaining('ghost'),
    });
  });
});

describe('TRANSFORMS — aggregates', () => {
  it('sum / avg / min / max', () => {
    expect(TRANSFORMS.sum!(SALES, 'amount')).toEqual({ kind: 'number', value: 700 });
    expect(TRANSFORMS.avg!(SALES, 'amount')).toEqual({ kind: 'number', value: 175 });
    expect(TRANSFORMS.min!(SALES, 'amount')).toEqual({ kind: 'number', value: 50 });
    expect(TRANSFORMS.max!(SALES, 'amount')).toEqual({ kind: 'number', value: 300 });
  });

  it('count → row count', () => {
    expect(TRANSFORMS.count!(SALES, '')).toEqual({ kind: 'number', value: 4 });
  });

  it('errors when an aggregate runs on a non-table (e.g. sum on a number)', () => {
    expect(TRANSFORMS.sum!({ kind: 'number', value: 1 }, 'amount')).toMatchObject({
      error: expect.stringContaining('table'),
    });
  });

  it('errors when the column has no numeric values', () => {
    expect(TRANSFORMS.sum!(SALES, 'region')).toMatchObject({ error: expect.any(String) });
  });
});

describe('TRANSFORMS — sort / head / tail', () => {
  it('sorts numerically asc and desc', () => {
    const asc = TRANSFORMS.sort!(SALES, 'amount') as Extract<Value, { kind: 'table' }>;
    expect(asc.rows.map((r) => r[1])).toEqual(['50', '100', '250', '300']);
    const desc = TRANSFORMS.sort!(SALES, 'amount desc') as Extract<Value, { kind: 'table' }>;
    expect(desc.rows.map((r) => r[1])).toEqual(['300', '250', '100', '50']);
  });

  it('head / tail take the first / last n rows', () => {
    expect(
      (TRANSFORMS.head!(SALES, '2') as Extract<Value, { kind: 'table' }>).rows.map((r) => r[0]),
    ).toEqual(['East', 'West']);
    expect(
      (TRANSFORMS.tail!(SALES, '1') as Extract<Value, { kind: 'table' }>).rows.map((r) => r[0]),
    ).toEqual(['North']);
  });

  it('head errors on a non-integer', () => {
    expect(TRANSFORMS.head!(SALES, 'x')).toMatchObject({ error: expect.any(String) });
  });
});

describe('TRANSFORMS — sed', () => {
  it('sed replaces the first match per row cell on a table (no /g flag)', () => {
    const t: Value = { kind: 'table', columns: ['Region'], rows: [['East Coast'], ['West Coast']] };
    const out = TRANSFORMS.sed!(t, 's/Coast/Region/');
    expect(out).toEqual({
      kind: 'table',
      columns: ['Region'],
      rows: [['East Region'], ['West Region']],
    });
  });

  it('sed with /g replaces every match in a cell', () => {
    const t: Value = { kind: 'table', columns: ['Label'], rows: [['aa-aa']] };
    expect(TRANSFORMS.sed!(t, 's/a/x/g')).toEqual({
      kind: 'table',
      columns: ['Label'],
      rows: [['xx-xx']],
    });
  });

  it('sed operates on a text Value directly', () => {
    expect(TRANSFORMS.sed!({ kind: 'text', value: 'hello world' }, 's/world/there/')).toEqual({
      kind: 'text',
      value: 'hello there',
    });
  });

  it('sed rejects a number Value', () => {
    const out = TRANSFORMS.sed!({ kind: 'number', value: 42 }, 's/4/9/');
    expect(out).toHaveProperty('error');
  });

  it('sed rejects a malformed s/// expression', () => {
    const out = TRANSFORMS.sed!({ kind: 'text', value: 'x' }, 'not-a-sed-expr');
    expect(out).toHaveProperty('error');
    expect((out as { error: string }).error).toMatch(/usage: sed/i);
  });
});

describe('TRANSFORMS — derive', () => {
  it('derive appends a computed column from two existing columns', () => {
    const table: Value = {
      kind: 'table',
      columns: ['Budget', 'Actual'],
      rows: [
        ['100', '80'],
        ['200', '250'],
      ],
    };
    const out = TRANSFORMS.derive!(table, 'Variance = Budget - Actual');
    expect(out).toEqual({
      kind: 'table',
      columns: ['Budget', 'Actual', 'Variance'],
      rows: [
        ['100', '80', '20'],
        ['200', '250', '-50'],
      ],
    });
  });

  it('derive supports a column and a literal operand', () => {
    const table: Value = { kind: 'table', columns: ['Revenue'], rows: [['100']] };
    expect(TRANSFORMS.derive!(table, 'Doubled = Revenue * 2')).toEqual({
      kind: 'table',
      columns: ['Revenue', 'Doubled'],
      rows: [['100', '200']],
    });
  });

  it('derive rejects a reference to an unknown column', () => {
    const table: Value = { kind: 'table', columns: ['A'], rows: [['1']] };
    const out = TRANSFORMS.derive!(table, 'X = A - B');
    expect(out).toHaveProperty('error');
    expect((out as { error: string }).error).toMatch(/unknown column.*B/i);
  });

  it('derive rejects a non-numeric cell for the row it fails on', () => {
    const table: Value = { kind: 'table', columns: ['A', 'B'], rows: [['x', '1']] };
    const out = TRANSFORMS.derive!(table, 'C = A + B');
    expect(out).toHaveProperty('error');
  });

  it('derive rejects a non-table Value', () => {
    expect(TRANSFORMS.derive!({ kind: 'number', value: 1 }, 'X = A + B')).toHaveProperty('error');
  });

  it('derive rejects a malformed expression', () => {
    expect(
      TRANSFORMS.derive!({ kind: 'table', columns: ['A'], rows: [] }, 'not an expr'),
    ).toHaveProperty('error');
  });
});

/* ───────────────────────────── evalExpr ───────────────────────────── */

/** A fake read host: maps a source to canned text (a GFM table for `read`). */
function fakeRunRead(map: Partial<Record<string, string>>): RunRead {
  return (source: Exclude<PipeSource, { src: 'var' }>): Promise<string> => {
    const key =
      source.src === 'read'
        ? `read:${source.selector}`
        : source.src === 'search'
          ? `search:${source.text}`
          : 'outline';
    return Promise.resolve(map[key] ?? map['*'] ?? '');
  };
}

const SALES_GFM = renderValue(SALES);

describe('evalExpr — end to end', () => {
  it('read | filter | sum → a number', async () => {
    const expr = parseExpressionLine('read Sales!A1:C9 | filter region=East | sum amount');
    const out = await evalExpr(expr as never, new Map(), fakeRunRead({ '*': SALES_GFM }));
    expect(out).toEqual({ kind: 'number', value: 150 });
  });

  it('let $t = read X binds, then $t | count reuses it across calls (shared env)', async () => {
    const env = new Map<string, Value>();
    const run = fakeRunRead({ '*': SALES_GFM });

    const bind = await evalExpr(parseExpressionLine('let $t = read Sales') as never, env, run);
    expect(isEvalError(bind)).toBe(false);
    expect(env.has('t')).toBe(true);

    const count = await evalExpr(parseExpressionLine('$t | count') as never, env, run);
    expect(count).toEqual({ kind: 'number', value: 4 });
  });

  it('an unbound $var is a corrective error', async () => {
    const out = await evalExpr(
      parseExpressionLine('$missing | count') as never,
      new Map(),
      fakeRunRead({}),
    );
    expect(out).toMatchObject({ error: expect.stringContaining('unbound') });
  });

  it('an unknown transform is corrective', async () => {
    const out = await evalExpr(
      parseExpressionLine('read X | frobnicate y') as never,
      new Map(),
      fakeRunRead({ '*': SALES_GFM }),
    );
    expect(out).toMatchObject({ error: expect.stringContaining('unknown transform') });
  });

  it('a non-table read degrades to a text Value', async () => {
    const out = await evalExpr(
      parseExpressionLine('read X') as never,
      new Map(),
      fakeRunRead({ '*': 'plain prose' }),
    );
    expect(out).toEqual({ kind: 'text', value: 'plain prose' });
  });

  it('a pipe-into-effect is rejected (Phase-1 pure-only) — the parser produces the corrective', () => {
    // `read X | set …` parses with a `set` STAGE; the runtime evaluator rejects it.
    const piped = parseExpressionLine('read X | set Sales!F2 =1');
    expect('error' in piped).toBe(false); // structurally a pipeline with a `set` stage
  });

  it('a pipe-into-effect stage is rejected at eval time', async () => {
    const out = await evalExpr(
      parseExpressionLine('read X | set Sales!F2 =1') as never,
      new Map(),
      fakeRunRead({ '*': SALES_GFM }),
    );
    expect(out).toMatchObject({ error: expect.stringContaining("can't be composed") });
  });
});

/* ───────────────────────────── C1 richer algebra ───────────────────────────── */

describe('C1 — filter boolean predicates', () => {
  const rows = (v: Value | { error: string }) =>
    'error' in v ? v : v.kind === 'table' ? v.rows.map((r) => r[2]) : v;

  it('keeps single-predicate syntax unchanged', () => {
    expect(rows(TRANSFORMS.filter!(SALES, 'region=East'))).toEqual(['Ann', 'Cy']);
  });

  it('and / or / not with precedence (not > and > or)', () => {
    expect(rows(TRANSFORMS.filter!(SALES, 'region=East and amount>=100'))).toEqual(['Ann']);
    expect(rows(TRANSFORMS.filter!(SALES, 'region=West or region=North'))).toEqual(['Bob', 'Dee']);
    expect(rows(TRANSFORMS.filter!(SALES, 'not region=East'))).toEqual(['Bob', 'Dee']);
    // `a or b and c` = `a or (b and c)`
    expect(rows(TRANSFORMS.filter!(SALES, 'region=West or region=East and amount<100'))).toEqual([
      'Bob',
      'Cy',
    ]);
  });

  it('parentheses override precedence', () => {
    expect(rows(TRANSFORMS.filter!(SALES, '(region=West or region=East) and amount<100'))).toEqual([
      'Cy',
    ]);
    expect(rows(TRANSFORMS.filter!(SALES, 'not (region=East or region=West)'))).toEqual(['Dee']);
  });

  it('quoted values keep keywords and spaces as data', () => {
    const t = table(['title'], [['rock and roll'], ['jazz']]);
    expect(TRANSFORMS.filter!(t, 'title contains "and roll"')).toMatchObject({
      rows: [['rock and roll']],
    });
    expect(TRANSFORMS.filter!(t, 'title = "rock and roll"')).toMatchObject({
      rows: [['rock and roll']],
    });
  });

  it('malformed predicates are corrective', () => {
    expect(TRANSFORMS.filter!(SALES, 'region=East and')).toMatchObject({
      error: expect.stringContaining('ends early'),
    });
    expect(TRANSFORMS.filter!(SALES, '(region=East')).toMatchObject({
      error: expect.stringContaining('missing )'),
    });
    expect(TRANSFORMS.filter!(SALES, 'region=East)')).toMatchObject({
      error: expect.stringContaining('unexpected ")"'),
    });
    expect(TRANSFORMS.filter!(SALES, 'nope=1 or region=East')).toMatchObject({
      error: expect.stringContaining('unknown column "nope"'),
    });
  });
});

describe('C1 — join / group / distinct / rename', () => {
  const REPS = table(
    ['rep', 'team', 'amount'],
    [
      ['Ann', 'A', '1'],
      ['Bob', 'B', '2'],
      ['Bob', 'B2', '3'],
    ],
  );
  const ctx = { env: new Map<string, Value>([['reps', REPS]]) };

  it('join is inner by default, one row per match, right key dropped, collisions suffixed', () => {
    const out = TRANSFORMS.join!(SALES, '$reps on rep', ctx);
    expect(out).toEqual(
      table(
        ['region', 'amount', 'rep', 'team', 'amount_2'],
        [
          ['East', '100', 'Ann', 'A', '1'],
          ['West', '250', 'Bob', 'B', '2'],
          ['West', '250', 'Bob', 'B2', '3'],
        ],
      ),
    );
  });

  it('join left keeps unmatched rows with blanks; `on a=b` maps differing key names', () => {
    const other = table(['name', 'team'], [['Ann', 'A']]);
    const out = TRANSFORMS.join!(SALES, '$o on rep=name left', { env: new Map([['o', other]]) });
    expect(out).toMatchObject({
      columns: ['region', 'amount', 'rep', 'team'],
      rows: [
        ['East', '100', 'Ann', 'A'],
        ['West', '250', 'Bob', ''],
        ['East', '50', 'Cy', ''],
        ['North', '300', 'Dee', ''],
      ],
    });
  });

  it('join errors: unbound, non-table, unknown columns, malformed', () => {
    expect(TRANSFORMS.join!(SALES, '$nope on rep', ctx)).toMatchObject({
      error: expect.stringContaining('unbound'),
    });
    expect(
      TRANSFORMS.join!(SALES, '$n on rep', { env: new Map([['n', { kind: 'number', value: 1 }]]) }),
    ).toMatchObject({
      error: expect.stringContaining('not a table'),
    });
    expect(TRANSFORMS.join!(SALES, '$reps on zzz', ctx)).toMatchObject({
      error: expect.stringContaining('unknown column "zzz"'),
    });
    expect(TRANSFORMS.join!(SALES, '$reps on rep=zzz', ctx)).toMatchObject({
      error: expect.stringContaining('"zzz" in $reps'),
    });
    expect(TRANSFORMS.join!(SALES, 'reps on rep', ctx)).toMatchObject({
      error: expect.stringContaining('join usage'),
    });
  });

  it('group aggregates per key in first-appearance order', () => {
    const out = TRANSFORMS.group!(
      SALES,
      'region agg sum(amount) as total, count(*) as n, avg(amount) as mean',
    );
    expect(out).toEqual(
      table(
        ['region', 'total', 'n', 'mean'],
        [
          ['East', '150', '2', '75'],
          ['West', '250', '1', '250'],
          ['North', '300', '1', '300'],
        ],
      ),
    );
  });

  it('group errors: unknown column, bad aggregate, duplicate output, no numeric values', () => {
    expect(TRANSFORMS.group!(SALES, 'zzz agg sum(amount) as t')).toMatchObject({
      error: expect.stringContaining('unknown column "zzz"'),
    });
    expect(TRANSFORMS.group!(SALES, 'region agg median(amount) as t')).toMatchObject({
      error: expect.stringContaining('malformed aggregate'),
    });
    expect(TRANSFORMS.group!(SALES, 'region agg sum(amount) as region')).toMatchObject({
      error: expect.stringContaining('duplicate output'),
    });
    expect(TRANSFORMS.group!(SALES, 'region agg sum(rep) as t')).toMatchObject({
      error: expect.stringContaining('no numeric values'),
    });
    expect(TRANSFORMS.group!(SALES, 'region')).toMatchObject({
      error: expect.stringContaining('group usage'),
    });
  });

  it('distinct over whole rows and over a projection', () => {
    expect(TRANSFORMS.distinct!(SALES, 'region')).toEqual(
      table(['region'], [['East'], ['West'], ['North']]),
    );
    const dup = table(['a'], [['1'], ['1'], ['2']]);
    expect(TRANSFORMS.distinct!(dup, '')).toEqual(table(['a'], [['1'], ['2']]));
    expect(TRANSFORMS.distinct!(SALES, 'zzz')).toMatchObject({
      error: expect.stringContaining('distinct: unknown column'),
    });
  });

  it('rename columns; refuses unknown sources and existing targets', () => {
    expect(TRANSFORMS.rename!(SALES, 'region=Region, amount=Total')).toMatchObject({
      columns: ['Region', 'Total', 'rep'],
    });
    expect(TRANSFORMS.rename!(SALES, 'zzz=a')).toMatchObject({
      error: expect.stringContaining('unknown column'),
    });
    expect(TRANSFORMS.rename!(SALES, 'region=rep')).toMatchObject({
      error: expect.stringContaining('already exists'),
    });
    expect(TRANSFORMS.rename!(SALES, 'region')).toMatchObject({
      error: expect.stringContaining('malformed pair'),
    });
  });
});

describe('C1 — derive functions', () => {
  const T = table(
    ['name', 'amount', 'start', 'end'],
    [
      [' ann ', '1.2345', '2026-01-01', '2026-01-31'],
      ['Bob', '2', '2026-03-15', '2026-03-15'],
    ],
  );
  const col = (v: Value | { error: string }) =>
    'error' in v ? v : v.kind === 'table' ? v.rows.map((r) => r.at(-1)) : v;

  it('text functions', () => {
    expect(col(TRANSFORMS.derive!(T, 'u = upper(name)'))).toEqual([' ANN ', 'BOB']);
    expect(col(TRANSFORMS.derive!(T, 'l = lower(name)'))).toEqual([' ann ', 'bob']);
    expect(col(TRANSFORMS.derive!(T, 't = trim(name)'))).toEqual(['ann', 'Bob']);
    expect(col(TRANSFORMS.derive!(T, 'n = len(name)'))).toEqual(['5', '3']);
    expect(col(TRANSFORMS.derive!(T, 'c = concat(name, " - ", amount)'))).toEqual([
      ' ann  - 1.2345',
      'Bob - 2',
    ]);
  });

  it('round and the date functions', () => {
    expect(col(TRANSFORMS.derive!(T, 'r = round(amount, 2)'))).toEqual(['1.23', '2']);
    expect(col(TRANSFORMS.derive!(T, 'r = round(amount)'))).toEqual(['1', '2']);
    expect(col(TRANSFORMS.derive!(T, 'y = year(start)'))).toEqual(['2026', '2026']);
    expect(col(TRANSFORMS.derive!(T, 'm = month(start)'))).toEqual(['1', '3']);
    expect(col(TRANSFORMS.derive!(T, 'd = day(end)'))).toEqual(['31', '15']);
    expect(col(TRANSFORMS.derive!(T, 'days = datediff(end, start)'))).toEqual(['30', '0']);
  });

  it('the arithmetic form is unchanged', () => {
    expect(col(TRANSFORMS.derive!(T, 'double = amount * 2'))).toEqual(['2.469', '4']);
  });

  it('errors: unknown function, arity, non-numeric, non-date, unknown arg', () => {
    expect(TRANSFORMS.derive!(T, 'x = median(amount)')).toMatchObject({
      error: expect.stringContaining('unknown function'),
    });
    expect(TRANSFORMS.derive!(T, 'x = datediff(start)')).toMatchObject({
      error: expect.stringContaining('takes 2'),
    });
    expect(TRANSFORMS.derive!(T, 'x = round(name)')).toMatchObject({
      error: expect.stringContaining('non-numeric'),
    });
    expect(TRANSFORMS.derive!(T, 'x = year(name)')).toMatchObject({
      error: expect.stringContaining('not a date'),
    });
    expect(TRANSFORMS.derive!(T, 'x = upper(zzz)')).toMatchObject({
      error: expect.stringContaining('unknown column or literal'),
    });
  });
});

describe('C1 — literal sources evaluate without a host read', () => {
  it('number / text / table literals and a $var join', async () => {
    const env = new Map<string, Value>();
    const run = fakeRunRead({});
    expect(await evalExpr(parseExpressionLine('let $n = 5') as never, env, run)).toEqual({
      kind: 'number',
      value: 5,
    });
    expect(await evalExpr(parseExpressionLine('let $s = "hi"') as never, env, run)).toEqual({
      kind: 'text',
      value: 'hi',
    });
    expect(
      await evalExpr(
        parseExpressionLine('let $t = table [["a","b"],["1","x"],["2","y"]]') as never,
        env,
        run,
      ),
    ).toEqual(
      table(
        ['a', 'b'],
        [
          ['1', 'x'],
          ['2', 'y'],
        ],
      ),
    );
    expect(
      await evalExpr(parseExpressionLine('$t | filter a>1 | count') as never, env, run),
    ).toEqual({
      kind: 'number',
      value: 1,
    });
    await evalExpr(parseExpressionLine('let $o = table [["a","c"],["1","Z"]]') as never, env, run);
    expect(
      await evalExpr(parseExpressionLine('$t | join $o on a left') as never, env, run),
    ).toEqual(
      table(
        ['a', 'b', 'c'],
        [
          ['1', 'x', 'Z'],
          ['2', 'y', ''],
        ],
      ),
    );
  });
});
