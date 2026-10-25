import { describe, expect, it } from 'vitest';
import { isUnsafeFormula } from './formula-policy.js';
import { hasScalarFormulaFootprint } from './formula-footprint.js';

describe('native formula output bounds', () => {
  it.each([
    '=A1*2',
    '=SUM(A1:A100)',
    "=SUM('Sales tab'!A:A)",
    '=IF(A1>0,ROUND(B1,2),0)',
    '=SUMIF(A1:A10,"East",B1:B10)',
    '=COUNTIFS(A1:A10,"East",B1:B10,">0")',
    '=VLOOKUP(A1,B1:D10,2,FALSE)',
    '=SUM(FILTER(A1:A10,B1:B10>0))',
    '=TEXTJOIN(",",TRUE,A1:A10)',
  ])('proves a scalar result for %s', (formula) =>
    expect(hasScalarFormulaFootprint(formula)).toBe(true),
  );
  it.each([
    '=A1:A10',
    '=SEQUENCE(1000)',
    '=FILTER(A1:A10,B1:B10>0)',
    '=IF(A1:A10>0,1,0)',
    '=SUMIF(A1:A10,{"East","West"},B1:B10)',
    '=SUMIFS(B1:B10,A1:A10,{"East","West"})',
    '=COUNTIF(A1:A10,{1,2})',
    '=VLOOKUP(A1,B1:D10,{2,3},FALSE)',
    '=NamedRange',
    '=MY_CUSTOM_FUNCTION(A1)',
    '=SUM(MY_CUSTOM_FUNCTION(A1))',
    '=LET(x,A1:A10,x)',
    '=A1#',
    '=Table1[Amount]',
  ])('stages an unknown or array footprint for %s', (formula) =>
    expect(hasScalarFormulaFootprint(formula)).toBe(false),
  );
  it('does not mistake references inside text for spill geometry', () => {
    expect(hasScalarFormulaFootprint('="A1:A100 SEQUENCE(99999)"')).toBe(true);
    expect(hasScalarFormulaFootprint('=IF(A1="say ""yes""",1,0)')).toBe(true);
  });
});

describe('external workbook reference guard', () => {
  it.each([
    "=SUM('[1]Sheet1'!A1)",
    '=SUM([Budget]Sheet1!A1)',
    "=SUM('[1]Sheet+1'!A1)",
    "=SUM('[1]Sheet=1'!A1)",
    "=SUM('[1]Sheet(1)'!A1)",
    "=SUM('[1]Sheet,1'!A1)",
    "=SUM('[1]Sheet''s data'!A1)",
    "=SUM('https://example.test/[Book]Sheet1'!A1)",
  ])('refuses %s', (formula) => expect(isUnsafeFormula(formula)).toBe(true));
  it('keeps workbook-looking text as inert string data', () => {
    expect(isUnsafeFormula('="[1]Sheet1!A1"')).toBe(false);
  });
});
