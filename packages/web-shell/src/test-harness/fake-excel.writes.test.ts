import { describe, it, expect, afterEach } from 'vitest';
import { ExcelBridge } from '@ge/bridge-excel';
import { ActuationRequestSchema, type ActuationRequest } from '@ge/contracts';
import { installFakeExcel, excelSeed, type ExcelSimulator } from './fake-excel.js';

/**
 * The REAL `ExcelBridge` object writes (table / chart / conditional format / comment) against the
 * Excel simulator: each lands `ok: true` and the workbook snapshot shows the change.
 */

let sim: ExcelSimulator | undefined;

afterEach(() => {
  sim?.restore();
  sim = undefined;
});

function install(): ExcelSimulator {
  const rows = [['Region', 'Q1', 'Q2', 'Revenue', 'Owner']];
  for (let i = 1; i <= 10; i++)
    rows.push([`R${i}`, String(i * 10), String(i * 20), String(i * 25000), `P${i}`]);
  sim = installFakeExcel(excelSeed({ sheets: [{ name: 'Sales', origin: 'A1', values: rows }] }));
  return sim;
}

function request(
  kind: string,
  params: Record<string, unknown>,
  id = `c-${kind}`,
): ActuationRequest {
  return ActuationRequestSchema.parse({ changeId: id, kind, surface: 'excel', params });
}

describe('fake-excel object writes driven by the real ExcelBridge', () => {
  it('binds mixed formula/literal readback to the same snapshot as a fresh capture', async () => {
    install();
    const bridge = new ExcelBridge();
    const result = await bridge.actuate(
      request('write-cells', {
        target: { range: 'Sales!G2:I2' },
        cellValues: [[null, '=literal', 7]],
        cellFormulas: [['=SUM(B2:C2)', '', '']],
      }),
    );
    const captured = await bridge.captureCells('Sales!G2:I2');
    expect(result.verification).toMatchObject({ status: 'verified', afterHash: captured.hash });
    expect(captured.values).toEqual([['=SUM(B2:C2)', '=literal', '7']]);
    expect(captured.formulas).toEqual([['=SUM(B2:C2)', '', '']]);
  });

  it('consumes a queued write once and refreshes an existing proxy on load', async () => {
    const s = install();
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getItem('Sales').getRange('G2');
      range.formulas = [['=SUM(B2:C2)']];
      range.values = [[null]];
      range.load('values,formulas');
      await ctx.sync();
      expect(range.values).toEqual([['=SUM(B2:C2)']]);
      // A user edits the workbook after the first sync. A later read may not replay the old write.
      s.seed.sheets[0]!.values[1]![6] = '99';
      range.load('values,formulas');
      await ctx.sync();
      expect(range.values).toEqual([['99']]);
      expect(range.formulas).toEqual([['99']]);
    });
    expect(s.snapshot().sheets[0]!.values[1]![6]).toBe('99');
  });

  it('moves formula identity with sorted rows while retaining formula-looking literal text', async () => {
    sim = installFakeExcel(
      excelSeed({
        sheets: [
          {
            name: 'Sales',
            origin: 'A1',
            values: [
              ['Key', 'Calc'],
              ['2', '=A2*2'],
              ['1', ''],
            ],
          },
        ],
      }),
    );
    const bridge = new ExcelBridge();
    await bridge.actuate(
      request('write-cells', { target: { range: 'Sales!B3' }, cellValues: [['=literal']] }),
    );
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getItem('Sales').getRange('A1:B3');
      range.sort.apply([{ key: 0, ascending: true }], false, true);
      await ctx.sync();
    });
    const captured = await bridge.captureCells('Sales!A1:B3');
    expect(captured.values).toEqual([
      ['Key', 'Calc'],
      ['1', '=literal'],
      ['2', '=A2*2'],
    ]);
    expect(captured.formulas).toEqual([
      ['', ''],
      ['', ''],
      ['', '=A2*2'],
    ]);
  });

  it('create-table records the table and returns the host-minted name as the inverse', async () => {
    const s = install();
    const res = await new ExcelBridge().actuate(
      request('create-table', { table: { range: 'Sales!A1:E11', hasHeaders: true } }),
    );
    expect(res).toMatchObject({
      ok: true,
      location: 'Sales!A1:E11',
      inverse: { op: 'delete-object', objectType: 'table', name: 'Table1' },
    });
    expect(s.snapshot().tables).toEqual([
      { name: 'Table1', range: 'Sales!A1:E11', hasHeaders: true },
    ]);
    // A second table gets the next minted name.
    const again = await new ExcelBridge().actuate(
      request('create-table', { table: { range: 'Sales!G1:H3', hasHeaders: true } }, 'c-2'),
    );
    expect(again.inverse).toMatchObject({ name: 'Table2' });
  });

  it('does not insert an extra header row outside the reviewed table footprint', async () => {
    const s = install();
    expect(
      await new ExcelBridge().actuate(
        request('create-table', { table: { range: 'Sales!G1:H3', hasHeaders: false } }),
      ),
    ).toMatchObject({ ok: false, error: { code: 'invalid_footprint' } });
    expect(s.snapshot().tables).toEqual([]);
  });

  it('insert-chart records type, source, title and the minted chart name', async () => {
    const s = install();
    const res = await new ExcelBridge().actuate(
      request('insert-chart', {
        chart: { chartType: 'column', sourceRange: 'Sales!A1:D11', title: 'Q3' },
      }),
    );
    expect(res).toMatchObject({
      ok: true,
      location: 'Chart 1',
      inverse: { op: 'delete-object', objectType: 'chart', name: 'Chart 1' },
    });
    expect(s.snapshot().charts).toEqual([
      {
        name: 'Chart 1',
        sheet: 'Sales',
        type: 'ColumnClustered',
        source: 'Sales!A1:D11',
        seriesBy: 'Auto',
        title: 'Q3',
        series: [{ values: 'Sales!A1:D11' }],
      },
    ]);
  });

  it('insert-chart over separate areas adds named series with shared categories', async () => {
    const s = install();
    const res = await new ExcelBridge().actuate(
      request('insert-chart', {
        chart: { chartType: 'line', sourceRange: 'Sales!A1:A11,Sales!B1:B11,Sales!C1:C11' },
      }),
    );
    expect(res.ok).toBe(true);
    const [chart] = s.snapshot().charts;
    expect(chart?.series).toEqual([
      { values: 'Sales!B1:B11', xValues: 'Sales!A2:A11' },
      { name: 'Q2', values: 'Sales!C2:C11', xValues: 'Sales!A2:A11' },
    ]);
  });

  it('format-conditional records a cellValue > 100000 rule with a green fill', async () => {
    const s = install();
    const res = await new ExcelBridge().actuate(
      request('format-conditional', {
        conditional: {
          range: 'Sales!D2:D11',
          rule: { kind: 'cellValue', operator: 'gt', value: '100000', fill: '#C6EFCE' },
        },
      }),
    );
    expect(res).toMatchObject({
      ok: true,
      location: 'Sales!D2:D11',
      inverse: { op: 'clear-conditional', range: 'Sales!D2:D11', ruleOrdinal: 0 },
    });
    expect(s.snapshot().conditionalFormats.get('Sales!D2:D11')).toEqual([
      {
        cfType: 'CellValue',
        cellValue: { operator: 'GreaterThan', formula1: '100000', fill: '#C6EFCE' },
      },
    ]);
  });

  it('add-comment anchors to the first cell and records the content', async () => {
    const s = install();
    const bridge = new ExcelBridge();
    const res = await bridge.actuate(
      request('add-comment', { target: { range: 'Sales!D5' }, text: 'why so low?' }),
    );
    expect(res).toMatchObject({ ok: true, location: 'Sales!D5' });
    await bridge.actuate(
      request('add-comment', { target: { range: 'Sales!B2:C2' }, text: 'check' }, 'c-2'),
    );
    expect(s.snapshot().comments).toEqual([
      {
        id: 'sim-comment-1',
        cell: 'Sales!D5',
        content: 'why so low?',
        replies: [],
        resolved: false,
      },
      { id: 'sim-comment-2', cell: 'Sales!B2', content: 'check', replies: [], resolved: false },
    ]);
    // comment-reply still resolves the added comment by id.
    const reply = await bridge.actuate(
      request('comment-reply', { target: { commentId: 'sim-comment-1' }, text: 'seasonal' }, 'c-3'),
    );
    expect(reply.ok).toBe(true);
    expect(s.snapshot().comments[0]?.replies).toEqual(['seasonal']);
  });
});
