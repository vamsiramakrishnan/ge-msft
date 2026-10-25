import { describe, expect, it } from 'vitest';
import { excelRowUnit, formulaDependencies } from './formula-dependencies.js';

const current = { id: 'sales-id', name: 'Sales' };
const sheets = [current, { id: 'other-id', name: 'Other Data' }];

describe('native formula evidence dependencies', () => {
  it('binds literal references to native worksheet IDs and absolute grid buckets', () => {
    const result = formulaDependencies(
      `=IF(A1="B2", 'Other Data'!$C$3:D4, LOG10(AF12))`,
      current,
      sheets,
    );
    expect(result.complete).toBe(true);
    expect(result.relatedIds).toEqual([
      'sheet:sales-id/row:1/col:1',
      'sheet:other-id/row:3/col:1',
      'sheet:other-id/row:4/col:1',
      'sheet:sales-id/row:12/col:1',
    ]);
    expect(excelRowUnit('native', 7, 33)).toBe('sheet:native/row:7/col:33');
  });
  it('handles quoted apostrophes and repeated ranges without duplicate links', () => {
    const special = { id: 'id', name: "Owner's Data" };
    expect(
      formulaDependencies(`='Owner''s Data'!A1+'Owner''s Data'!B1`, current, [...sheets, special])
        .relatedIds,
    ).toEqual(['sheet:id/row:1/col:1']);
  });
  it('resolves native named ranges and conservatively covers structured table references', () => {
    const result = formulaDependencies(
      '=Revenue+SUM(SalesTable[[#Totals],[Amount]])',
      current,
      sheets,
      100,
      [
        { name: 'Revenue', address: "='Other Data'!A3:A4", kind: 'named-range' },
        { name: 'Sales!Revenue', address: '=Sales!AF2', kind: 'named-range' },
        { name: 'SalesTable', address: 'Sales!A7:AG8', kind: 'table' },
      ],
    );
    expect(result.complete).toBe(true);
    expect(result.relatedIds).toEqual([
      'sheet:sales-id/row:2/col:1',
      'sheet:sales-id/row:7/col:1',
      'sheet:sales-id/row:7/col:33',
      'sheet:sales-id/row:8/col:1',
      'sheet:sales-id/row:8/col:33',
    ]);
  });
  it('bounds repeated alias expansion before parsing references', () => {
    const result = formulaDependencies(
      '=' + Array.from({ length: 1000 }, () => 'Revenue').join('+'),
      current,
      sheets,
      100,
      [{ name: 'Revenue', address: 'A1:' + ' '.repeat(4000) + 'A1', kind: 'named-range' }],
    );
    expect(result.complete).toBe(false);
    expect(result.relatedIds.length).toBeLessThanOrEqual(100);
  });
  it.each([
    '=INDIRECT("A1")',
    '=OFFSET(A1, 1, 0)',
    '=Revenue+1',
    '=SUM(Table1[Amount])',
    "='[Budget.xlsx]Other Data'!A1",
    '=SUM(Sales:Other!A1)',
    '=SUM(A:A)',
    '=A1048577',
    '=XFE1',
  ])('marks unsupported references incomplete: %s', (formula) => {
    const result = formulaDependencies(formula, current, sheets);
    expect(result.complete).toBe(false);
    expect(result.limitations.join(' ')).toContain('require a fresh explicit read');
    expect(result.relatedIds.length).toBeLessThanOrEqual(100);
  });
  it('does not infer dependencies from string content or a formula-looking literal', () => {
    expect(formulaDependencies('="A1 and Other!B2"', current, sheets).relatedIds).toEqual([]);
    expect(formulaDependencies('ordinary A1', current, sheets).complete).toBe(false);
  });
  it('stops enormous range expansion at the link budget and reports omissions', () => {
    const result = formulaDependencies('=SUM(A1:XFD1048576)', current, sheets);
    expect(result.relatedIds).toHaveLength(100);
    expect(result.complete).toBe(false);
    expect(formulaDependencies('=' + 'A'.repeat(17 * 1024), current, sheets).relatedIds).toEqual(
      [],
    );
    expect(() => formulaDependencies('=A1', current, sheets, 101)).toThrow('budget');
  });
});
