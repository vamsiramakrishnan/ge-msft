import { describe, it, expect } from 'vitest';
import { asChangeId, type ActuationRequest } from '@ge/contracts';
import {
  planAddComment,
  planConditional,
  planCreateTable,
  planFormatCells,
  planInsertChart,
  planWriteCells,
} from './actuate-plan.js';

function req(params: ActuationRequest['params'], kind: ActuationRequest['kind']): ActuationRequest {
  return { changeId: asChangeId('c1'), kind, surface: 'excel', params };
}

describe('planWriteCells (shared grid precedence)', () => {
  it('prefers typed cells over a conflicting legacy grid and preserves scalar types', () => {
    expect(
      planWriteCells(
        req(
          {
            target: { range: 'Sales!B2' },
            cells: [['legacy']],
            cellValues: [
              [17, true],
              [null, '=literal'],
            ],
          },
          'write-cells',
        ),
      ),
    ).toEqual({
      address: 'Sales!B2',
      values: [
        [17, true],
        [null, '=literal'],
      ],
    });
  });

  it.each([
    { cells: [], cellValues: [[42]], expected: [[42]] },
    { cells: [['legacy']], cellValues: [], expected: [] },
    { cells: [['legacy']], expected: [['legacy']] },
    { expected: [] },
  ])('preserves empty-grid precedence: %j', ({ expected, ...params }) => {
    expect(planWriteCells(req(params, 'write-cells')).values).toEqual(expected);
  });
});

describe('planFormatCells (ADR-0004 format-cells)', () => {
  it('maps each present format facet to a host op and flags hasOps', () => {
    const plan = planFormatCells(
      req(
        {
          target: { range: 'Sales!A1:C1' },
          format: { bold: true, italic: false, fill: '#FFF2CC', numberFormat: '$#,##0.00' },
        },
        'format-cells',
      ),
    );
    expect(plan).toEqual({
      address: 'Sales!A1:C1',
      bold: true,
      italic: false,
      fill: '#FFF2CC',
      numberFormat: '$#,##0.00',
      hasOps: true,
    });
  });

  it('omits absent facets (undefined ⇒ leave untouched) and still reports hasOps for one facet', () => {
    const plan = planFormatCells(
      req({ target: { range: 'A1' }, format: { bold: true } }, 'format-cells'),
    );
    expect(plan).toEqual({ address: 'A1', bold: true, hasOps: true });
    expect(plan.italic).toBeUndefined();
    expect(plan.fill).toBeUndefined();
    expect(plan.numberFormat).toBeUndefined();
  });

  it('reports hasOps:false for an empty/absent format (a no-op the bridge degrades)', () => {
    expect(planFormatCells(req({ target: { range: 'A1' } }, 'format-cells')).hasOps).toBe(false);
    expect(
      planFormatCells(req({ target: { range: 'A1' }, format: {} }, 'format-cells')).hasOps,
    ).toBe(false);
  });

  it('omits the address when no target.range is given', () => {
    const plan = planFormatCells(req({ format: { bold: true } }, 'format-cells'));
    expect(plan.address).toBeUndefined();
    expect(plan.hasOps).toBe(true);
  });

  it('preserves an explicit false (e.g. unbold) rather than dropping it', () => {
    const plan = planFormatCells(
      req({ target: { range: 'A1' }, format: { bold: false } }, 'format-cells'),
    );
    expect(plan.bold).toBe(false);
    expect(plan.hasOps).toBe(true);
  });
});

describe('planAddComment (ADR-0004 add-comment, Excel)', () => {
  it('shapes the comment to the anchor range and single-lines the text', () => {
    const plan = planAddComment(
      req({ target: { range: 'Sheet1!B2' }, text: 'Spike   here\nplease' }, 'add-comment'),
    );
    expect(plan).toEqual({
      address: 'Sheet1!B2',
      text: 'Spike here please',
      hasTarget: true,
      hasText: true,
    });
  });

  it('reports hasTarget:false when no range is given', () => {
    const plan = planAddComment(req({ text: 'note' }, 'add-comment'));
    expect(plan.hasTarget).toBe(false);
    expect(plan.address).toBeUndefined();
  });

  it('reports hasText:false for empty/whitespace text', () => {
    expect(planAddComment(req({ target: { range: 'A1' } }, 'add-comment')).hasText).toBe(false);
    expect(
      planAddComment(req({ target: { range: 'A1' }, text: '   ' }, 'add-comment')).hasText,
    ).toBe(false);
  });
});

describe('planCreateTable (ADR-0007 table verb)', () => {
  it('extracts the range, hasHeaders and the requested name; flags hasTable', () => {
    const plan = planCreateTable(
      req({ table: { range: 'Sales!A1:C4', hasHeaders: false, name: 'Q1' } }, 'create-table'),
    );
    expect(plan).toEqual({
      address: 'Sales!A1:C4',
      hasHeaders: false,
      requestedName: 'Q1',
      hasTable: true,
    });
  });

  it('defaults hasHeaders to true when omitted', () => {
    const plan = planCreateTable(req({ table: { range: 'A1:C4' } as never }, 'create-table'));
    expect(plan.hasHeaders).toBe(true);
  });

  it('fails closed (hasTable:false, no address) when table params are absent', () => {
    const plan = planCreateTable(req({}, 'create-table'));
    expect(plan).toEqual({ hasHeaders: true, hasTable: false });
    expect(plan.address).toBeUndefined();
  });
});

describe('planInsertChart (ADR-0007 chart verb)', () => {
  it('maps each agent chart type to its Excel.ChartType string', () => {
    const map: Array<[string, string]> = [
      ['column', 'ColumnClustered'],
      ['bar', 'BarClustered'],
      ['line', 'Line'],
      ['pie', 'Pie'],
      ['scatter', 'XYScatter'],
      ['area', 'Area'],
    ];
    for (const [agent, host] of map) {
      const plan = planInsertChart(
        req(
          { chart: { chartType: agent as 'column', sourceRange: 'A1:C4', seriesBy: 'auto' } },
          'insert-chart',
        ),
      );
      expect(plan.chartType).toBe(host);
    }
  });

  it('maps seriesBy and carries the title; flags hasChart', () => {
    const plan = planInsertChart(
      req(
        {
          chart: { chartType: 'line', sourceRange: 'Sales!A1:C4', seriesBy: 'rows', title: 'Rev' },
        },
        'insert-chart',
      ),
    );
    expect(plan).toEqual({
      address: 'Sales!A1:C4',
      chartType: 'Line',
      seriesBy: 'Rows',
      title: 'Rev',
      hasChart: true,
    });
  });

  it('fails closed (hasChart:false) when chart params are absent', () => {
    const plan = planInsertChart(req({}, 'insert-chart'));
    expect(plan.hasChart).toBe(false);
    expect(plan.address).toBeUndefined();
  });
});

describe('planConditional (ADR-0007 cf verb)', () => {
  it('maps a cellValue rule operator and carries fill/formula2', () => {
    const plan = planConditional(
      req(
        {
          conditional: {
            range: 'Sales!C2:C4',
            rule: {
              kind: 'cellValue',
              operator: 'between',
              value: '100',
              value2: '300',
              fill: '#C6EFCE',
            },
          },
        },
        'format-conditional',
      ),
    );
    expect(plan.hasConditional).toBe(true);
    expect(plan.address).toBe('Sales!C2:C4');
    expect(plan.rule).toEqual({
      kind: 'cellValue',
      cfType: 'CellValue',
      operator: 'Between',
      formula1: '100',
      formula2: '300',
      fill: '#C6EFCE',
    });
  });

  it('maps each cellValue operator', () => {
    const ops: Array<[string, string]> = [
      ['gt', 'GreaterThan'],
      ['lt', 'LessThan'],
      ['ge', 'GreaterThanOrEqual'],
      ['le', 'LessThanOrEqual'],
      ['eq', 'EqualTo'],
      ['ne', 'NotEqualTo'],
    ];
    for (const [agent, host] of ops) {
      const plan = planConditional(
        req(
          {
            conditional: {
              range: 'A1',
              rule: { kind: 'cellValue', operator: agent as 'gt', value: '1' },
            },
          },
          'format-conditional',
        ),
      );
      expect(plan.rule).toMatchObject({ operator: host });
    }
  });

  it('maps a top rule to TopItems / BottomItems by the bottom flag', () => {
    const top = planConditional(
      req(
        { conditional: { range: 'A1', rule: { kind: 'top', rank: 3, bottom: false } } },
        'format-conditional',
      ),
    );
    expect(top.rule).toEqual({ kind: 'top', cfType: 'TopBottom', rank: 3, criterion: 'TopItems' });
    const bottom = planConditional(
      req(
        { conditional: { range: 'A1', rule: { kind: 'top', rank: 3, bottom: true } } },
        'format-conditional',
      ),
    );
    expect(bottom.rule).toMatchObject({ criterion: 'BottomItems' });
  });

  it('maps bare dataBar / colorScale rules', () => {
    expect(
      planConditional(
        req({ conditional: { range: 'A1', rule: { kind: 'dataBar' } } }, 'format-conditional'),
      ).rule,
    ).toEqual({ kind: 'dataBar', cfType: 'DataBar' });
    expect(
      planConditional(
        req({ conditional: { range: 'A1', rule: { kind: 'colorScale' } } }, 'format-conditional'),
      ).rule,
    ).toEqual({ kind: 'colorScale', cfType: 'ColorScale' });
  });

  it('fails closed (hasConditional:false) when conditional params are absent', () => {
    const plan = planConditional(req({}, 'format-conditional'));
    expect(plan.hasConditional).toBe(false);
    expect(plan.rule).toBeUndefined();
  });

  it('flags an unsafe cellValue formula (untrusted active content) so the bridge degrades', () => {
    // A WEBSERVICE/DDE/external-ref payload as the threshold value must trip the formula screen —
    // a leading `=` is not required (Excel evaluates a CF formula by content).
    const web = planConditional(
      req(
        {
          conditional: {
            range: 'A1',
            rule: { kind: 'cellValue', operator: 'gt', value: 'WEBSERVICE("http://x")' },
          },
        },
        'format-conditional',
      ),
    );
    expect(web.unsafe).toBe(true);
    const dde = planConditional(
      req(
        {
          conditional: {
            range: 'A1',
            rule: { kind: 'cellValue', operator: 'gt', value: "cmd|'/c calc'!A1" },
          },
        },
        'format-conditional',
      ),
    );
    expect(dde.unsafe).toBe(true);
    // a normal numeric threshold is not flagged
    const ok = planConditional(
      req(
        { conditional: { range: 'A1', rule: { kind: 'cellValue', operator: 'gt', value: '100' } } },
        'format-conditional',
      ),
    );
    expect(ok.unsafe).toBeUndefined();
  });
});
