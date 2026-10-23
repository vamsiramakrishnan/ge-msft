import { describe, it, expect } from 'vitest';
import { quoteSheetNames } from './command-grammar.js';

describe('quoteSheetNames', () => {
  const sheets = ['Branch Performance', 'Sales', "O'Brien Q1"];

  it('quotes a known sheet name with a space written bare before !', () => {
    expect(quoteSheetNames('/sort-range range=Branch Performance!A1:I25 key=x', sheets)).toBe(
      "/sort-range range='Branch Performance'!A1:I25 key=x",
    );
    expect(quoteSheetNames('set D2 =SUM(Branch Performance!F2:F25)', sheets)).toBe(
      "set D2 =SUM('Branch Performance'!F2:F25)",
    );
    expect(quoteSheetNames("set O'Brien Q1!A1 1", sheets)).toBe("set 'O''Brien Q1'!A1 1");
  });

  it('leaves quoted names, simple names, unknown names and quoted values alone', () => {
    for (const line of [
      "set 'Branch Performance'!A1 1",
      'set Sales!A1 1',
      'set Other Sheet!A1 1',
      'set Sales!A1 "see Branch Performance!A1"',
      'set XBranch Performance!A1 1',
    ])
      expect(quoteSheetNames(line, sheets)).toBe(line);
  });

  it('never rewrites a sheet name that could pass for command text', () => {
    const line = '/write-cells range=Sales!B2 1';
    expect(quoteSheetNames(line, ['range=Sales', 'Sales'])).toBe(line);
    expect(quoteSheetNames('set A"B C!A1 1', ['A"B C'])).toBe('set A"B C!A1 1');
  });

  it('keeps a broken quote on one line from changing the next line', () => {
    expect(quoteSheetNames('set A1 "C:\\"\nset Branch Performance!A1 1', sheets)).toBe(
      'set A1 "C:\\"\nset \'Branch Performance\'!A1 1',
    );
  });
});
