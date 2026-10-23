import { describe, it, expect } from 'vitest';
import { canonicalSheetNames } from './command-grammar.js';

describe('canonicalSheetNames — a sheet name written without its stray spaces', () => {
  const sheets = ['Branch Performance', ' Ringkasan Cabang', 'Notes '];

  it('uses the real name when exactly one sheet matches without leading/trailing spaces', () => {
    expect(canonicalSheetNames('grid \'Ringkasan Cabang\'!A1:E5 = "x"', sheets)).toBe(
      'grid \' Ringkasan Cabang\'!A1:E5 = "x"',
    );
    expect(canonicalSheetNames('set Notes!A1 1', sheets)).toBe("set 'Notes '!A1 1");
    expect(canonicalSheetNames("set B2 =SUM('ringkasan cabang'!B2:B5)", sheets)).toBe(
      "set B2 =SUM(' Ringkasan Cabang'!B2:B5)",
    );
  });

  it('leaves existing names, unknown names, quoted values and ambiguous matches alone', () => {
    for (const line of [
      "set 'Branch Performance'!A1 1",
      "set ' Ringkasan Cabang'!A1 1",
      "set 'Other'!A1 1",
      'set A1 "see \'Ringkasan Cabang\'!A1"',
    ])
      expect(canonicalSheetNames(line, sheets)).toBe(line);
    expect(canonicalSheetNames("set 'Data'!A1 1", [' Data', 'Data '])).toBe("set 'Data'!A1 1");
    expect(canonicalSheetNames("set 'a=b'!A1 1", [' a=b'])).toBe("set 'a=b'!A1 1");
  });

  it('never redirects a name that is a real sheet, listed or not (security review H1)', () => {
    // The full live list includes a VeryHidden or 31st sheet `Sales`: it stays the target.
    expect(canonicalSheetNames("set 'Sales'!B2 1", ['Sales', ' Sales'])).toBe("set 'Sales'!B2 1");
    expect(canonicalSheetNames("set 'sales'!B2 1", ['SALES', ' Sales'])).toBe("set 'sales'!B2 1");
  });

  it('treats only ASCII spaces as stray', () => {
    expect(canonicalSheetNames("set 'Sales'!B2 1", ['\u00a0Sales'])).toBe("set 'Sales'!B2 1");
  });

  it('never reads the last word of a longer bare name as another sheet (security review M2)', () => {
    expect(canonicalSheetNames('set D2 =SUM(Data Cabang!F2:F25)', ['Data Cabang', 'Cabang '])).toBe(
      'set D2 =SUM(Data Cabang!F2:F25)',
    );
  });
});
