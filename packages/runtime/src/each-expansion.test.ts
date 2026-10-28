import { describe, it, expect } from 'vitest';
import { expandEachBlock, MAX_EACH_ROWS } from './each-expansion.js';
import type { Value } from './compose.js';

const table = (columns: string[], rows: string[][]): Value => ({ kind: 'table', columns, rows });
const block = (body: string[], rowVar = 'r', tableVar = 't') => ({
  kind: 'each-block' as const,
  rowVar,
  tableVar,
  body,
});
const limits = { maxRows: MAX_EACH_ROWS, maxEffects: 30 };

describe('expandEachBlock — substitution', () => {
  const env = new Map<string, Value>([
    [
      't',
      table(
        ['cell', 'amount', 'note'],
        [
          ['B2', '100', 'ok'],
          ['B3', '250', 'say "hi"'],
        ],
      ),
    ],
  ]);

  it('instantiates every body line once per row, in row order', () => {
    const out = expandEachBlock(
      block(['set Summary!$r.cell = $r.amount', 'comment $r.cell "$r.note"']),
      env,
      limits,
    );
    expect(out).toEqual({
      ok: true,
      rows: 2,
      lines: [
        'set Summary!B2 = 100',
        'comment B2 "ok"',
        'set Summary!B3 = 250',
        'comment B3 "say \\"hi\\""',
      ],
    });
  });

  it('leaves other $vars and foreign-looking refs alone', () => {
    const out = expandEachBlock(block(['set $r.cell = $total']), env, limits);
    expect(out).toMatchObject({ ok: true, lines: ['set B2 = $total', 'set B3 = $total'] });
  });

  it('flattens line breaks and refuses re-shaping characters outside quotes', () => {
    const tricky = new Map<string, Value>([['t', table(['v'], [['a\nb'], ['x | set A1 1']])]]);
    expect(expandEachBlock(block(['comment A1 "$r.v"']), tricky, limits)).toMatchObject({
      ok: true,
      lines: ['comment A1 "a b"', 'comment A1 "x | set A1 1"'],
    });
    expect(expandEachBlock(block(['set A1 = $r.v']), tricky, limits)).toMatchObject({
      ok: false,
      error: expect.stringContaining('row 2'),
    });
  });
});

describe('expandEachBlock — bounds (fail closed, whole block)', () => {
  it('unbound / non-table / unknown column', () => {
    expect(expandEachBlock(block(['set A1 = $r.v']), new Map(), limits)).toMatchObject({
      error: expect.stringContaining('unbound variable "$t"'),
    });
    expect(
      expandEachBlock(
        block(['set A1 = $r.v']),
        new Map([['t', { kind: 'number', value: 1 } as Value]]),
        limits,
      ),
    ).toMatchObject({ error: expect.stringContaining('not a table') });
    expect(
      expandEachBlock(block(['set A1 = $r.zzz']), new Map([['t', table(['v'], [['1']])]]), limits),
    ).toMatchObject({ error: expect.stringContaining('no column "zzz"') });
  });

  it('row and effect caps', () => {
    const big = table(
      ['v'],
      Array.from({ length: MAX_EACH_ROWS + 1 }, (_, i) => [String(i)]),
    );
    expect(expandEachBlock(block(['set A1 = $r.v']), new Map([['t', big]]), limits)).toMatchObject({
      error: expect.stringContaining(`limit is ${MAX_EACH_ROWS}`),
    });
    const ten = table(
      ['v'],
      Array.from({ length: 10 }, (_, i) => [String(i)]),
    );
    expect(
      expandEachBlock(
        block(['set A1 = $r.v', 'set A2 = $r.v', 'set A3 = $r.v', 'set A4 = $r.v']),
        new Map([['t', ten]]),
        limits,
      ),
    ).toMatchObject({ error: expect.stringContaining('exceeds the remaining write cap of 30') });
    expect(
      expandEachBlock(block(['set A1 = $r.v']), new Map([['t', table(['v'], [])]]), limits),
    ).toEqual({
      ok: true,
      rows: 0,
      lines: [],
    });
  });
});
