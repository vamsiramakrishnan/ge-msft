import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema } from '@ge/contracts';
import { rangeToContext, selectionValuesToContext, splitHeaderRows } from './capture.js';
import { planWriteCells } from './actuate-plan.js';

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

describe('excel actuation planning (pure)', () => {
  it('extracts address + values for write-cells', () => {
    const plan = planWriteCells({
      changeId: 'c1',
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
      changeId: 'c2',
      kind: 'write-cells',
      surface: 'excel',
      params: {},
    });
    expect(plan).toEqual({ values: [] });
  });
});
