import { describe, it, expect } from 'vitest';
import { formulasStayInRow } from './actuate-plan.js';

describe('formulasStayInRow — formulas a sort can move with their row', () => {
  // One formula in row 2 of a range spanning columns A:F.
  const ok = (formula: string) => formulasStayInRow([['', '', '', '', '', formula]], 2, 'A', false);

  it('allows built-in formulas that use only cells of their own row inside the range', () => {
    for (const f of [
      '=E2/D2',
      '=$E2/$D2',
      '=ROUND(E2/D2,2)',
      '=IF(A2="x 9",B2*2,0)',
      '=SUM(C2:F2)',
      '=LOG10(B2)',
      '=B2*1E5',
      '=TRUE',
      '=B2&" (Rp)"',
      '=IFERROR(E2/D2,"")',
    ])
      expect(ok(f), f).toBe(true);
    expect(
      formulasStayInRow(
        [
          ['', '=B2*2'],
          ['', '=B3*2'],
        ],
        2,
        'A',
        false,
      ),
    ).toBe(true);
  });

  it('refuses anything that points outside its row or range, or depends on position', () => {
    for (const f of [
      '=E3/D2',
      '=SUM(C2:C7)',
      '=E$2',
      '=$E$2',
      '=Z2*2',
      '=Sheet2!A2',
      "='Q1 Sales'!A2",
      '=[@Deposits]',
      '=SUM(A:A)',
      '=SUM(2:2)',
      '=SUM(B2:.B9)',
      '=SUM(A2#)',
      '=Rate*E2',
      '=\\rate*B2',
      '=税率*B2',
      '=XFE2*2',
      '=OFFSET(A2,0,1)',
      '=_xlfn.OFFSET(A2,0,1)',
      '=MYFN(B2)',
      '=INDIRECT("A2")',
      '=ROW()',
      '=INDEX(A2:C2,1)',
      '=RAND()*A2',
      '=@A2:A9',
    ])
      expect(ok(f), f).toBe(false);
  });

  it('skips the header row when the range has headers', () => {
    expect(formulasStayInRow([['=SUM(A2:A9)'], ['=A2*2']], 1, 'A', true)).toBe(true);
    expect(formulasStayInRow([['=SUM(A2:A9)'], ['=A2*2']], 1, 'A', false)).toBe(false);
  });
});
