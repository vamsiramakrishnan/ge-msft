import { describe, expect, it } from 'vitest';
import { groupedChartPlan } from './grouped-chart.js';

const data = [
  ['Branch', 'Region', 'Month', 'Complaints'],
  ['Jakarta', 'West', 'Jan', 4],
  ['Jakarta', 'West', 'Feb', 5],
  ['Medan', 'North', 'Jan', 6],
  ['Medan', 'North', 'Feb', 7],
];

describe('groupedChartPlan', () => {
  it('builds a Month × Branch summary of live exact-match sums next to the data (UAT #11)', () => {
    const plan = groupedChartPlan("'Branch Performance'!A1:D5", data, {
      x: 'month',
      by: 'Branch',
      value: 'Complaints',
    });
    expect(plan).toEqual({
      range: "'Branch Performance'!F1:H3",
      cells: [
        ['Month', 'Jakarta', 'Medan'],
        [
          'Jan',
          '=SUMPRODUCT(--($A$2:$A$5=G$1),--($C$2:$C$5=$F2),$D$2:$D$5)',
          '=SUMPRODUCT(--($A$2:$A$5=H$1),--($C$2:$C$5=$F2),$D$2:$D$5)',
        ],
        [
          'Feb',
          '=SUMPRODUCT(--($A$2:$A$5=G$1),--($C$2:$C$5=$F3),$D$2:$D$5)',
          '=SUMPRODUCT(--($A$2:$A$5=H$1),--($C$2:$C$5=$F3),$D$2:$D$5)',
        ],
      ],
    });
  });

  it('accepts column letters and moves right by an offset when the first spot is taken', () => {
    const plan = groupedChartPlan('Sheet1!A1:D5', data, { x: 'C', by: 'A', value: 'D' }, 4);
    expect(plan).toMatchObject({ range: 'Sheet1!J1:L3' });
  });

  it('writes a label that looks like a formula as text', () => {
    const rows = [
      ['Branch', 'Month', 'N'],
      ['=HYPERLINK("http://x","y")', 'Jan', 1],
    ];
    const plan = groupedChartPlan('S!A1:C2', rows, { x: 'Month', by: 'Branch', value: 'N' });
    expect(plan).toMatchObject({
      cells: [
        ['Month', `'=HYPERLINK("http://x","y")`],
        ['Jan', expect.any(String)],
      ],
    });
  });

  it('keeps one column per group as Excel compares labels (case-insensitive)', () => {
    const rows = [
      ['Branch', 'Month', 'N'],
      ['Jakarta', 'Jan', 1],
      ['JAKARTA', 'Jan', 2],
    ];
    const plan = groupedChartPlan('S!A1:C3', rows, { x: 'Month', by: 'Branch', value: 'N' });
    expect(plan).toMatchObject({
      cells: [
        ['Month', 'Jakarta'],
        ['Jan', expect.any(String)],
      ],
    });
  });

  it('names the headers when a column is not found', () => {
    expect(
      groupedChartPlan('S!A1:D5', data, { x: 'Bulan', by: 'Branch', value: 'Complaints' }),
    ).toEqual({
      error: expect.stringContaining('the headers are Branch, Region, Month, Complaints'),
    });
  });
});
