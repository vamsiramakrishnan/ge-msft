import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { asChangeId, makeCellSnapshot, type ActuationRequest, type CellValue } from '@ge/contracts';
import {
  applyNativeValidation,
  applyNativeFilter,
  applyNativePivot,
  applyNativeWorksheet,
  nativeRange,
} from './native-operations.js';
import { ExcelBridge } from './excel-bridge.js';

const state = vi.hoisted(() => ({ context: undefined as unknown, failSync: 0, syncs: 0 }));
vi.mock('./native-run.js', () => ({
  // Mirror the overload: `nativeRun(work)` or `nativeRun(options, work)`.
  nativeRun: (
    optionsOrWork: unknown | ((context: unknown) => unknown),
    maybeWork?: (context: unknown) => unknown,
  ) => (typeof optionsOrWork === 'function' ? optionsOrWork : maybeWork!)(state.context),
}));
vi.mock('./cell-snapshot.js', async (importOriginal) => {
  const original = await importOriginal<typeof import('./cell-snapshot.js')>();
  const { makeCellSnapshot } = await import('@ge/contracts');
  return {
    ...original,
    excelDocumentId: async () => 'workbook-id',
    snapshotRange: async (range: FakeRange) =>
      makeCellSnapshot({
        surface: 'excel',
        documentId: 'workbook-id',
        locator: range.address,
        objectId: range.worksheet.id,
        values: range.values,
        formulas: range.values.map((row) => row.map(() => '')),
      }),
  };
});

interface FakeRange {
  address: string;
  values: CellValue[][];
  rowCount: number;
  columnCount: number;
  worksheet: { id: string };
  dataValidation: {
    type: string;
    rule: Excel.DataValidationRule;
    errorAlert: Excel.DataValidationErrorAlert;
    load(): void;
  };
  load: () => void;
  getCell: () => FakeRange;
}
function range(address: string, values: CellValue[][], id = 'source-sheet'): FakeRange {
  const bounds = nativeRange(address)!;
  const validation = {
    type: 'None',
    currentRule: {} as Excel.DataValidationRule,
    errorAlert: {
      showAlert: false,
      style: 'Stop',
      title: '',
      message: '',
    } as Excel.DataValidationErrorAlert,
    load() {},
    get rule() {
      return validation.currentRule;
    },
    set rule(value: Excel.DataValidationRule) {
      validation.currentRule = value;
      validation.type = 'List';
    },
  };
  const item: FakeRange = {
    address,
    values,
    rowCount: bounds.rows,
    columnCount: bounds.columns,
    worksheet: { id },
    dataValidation: validation,
    load() {},
    getCell() {
      return item;
    },
  };
  return item;
}
async function hash(item: FakeRange): Promise<string> {
  return (
    await makeCellSnapshot({
      surface: 'excel',
      documentId: 'workbook-id',
      locator: item.address,
      objectId: item.worksheet.id,
      values: item.values,
      formulas: item.values.map((row) => row.map(() => '')),
    })
  ).hash;
}
function request(
  kind: ActuationRequest['kind'],
  params: ActuationRequest['params'],
): ActuationRequest {
  return { surface: 'excel', kind, params, changeId: asChangeId('native-operation-test') };
}
function host() {
  const source = range('Sales!A1:C4', [
    ['Region', 'Revenue', 'Cost'],
    ['East', 10, 2],
    ['West', 7, 3],
    ['East', 5, 4],
  ]);
  const target = range(
    'Report!A1:D7',
    Array.from({ length: 7 }, () => ['', '', '', '']),
    'target-sheet',
  );
  const output = range(
    'Report!A1:C3',
    [
      ['Region', 'Sum of Revenue', 'Sum of Cost'],
      ['East', 15, 6],
      ['West', 7, 3],
    ],
    'target-sheet',
  );
  const data = range(
    'Report!B2:C3',
    [
      [15, 6],
      [7, 3],
    ],
    'target-sheet',
  );
  const labels = range('Report!A1:A3', [['Region'], ['East'], ['West']], 'target-sheet');
  const filter = {
    enabled: false,
    criteria: [] as Excel.FilterCriteria[],
    load() {},
    apply: vi.fn((_range: unknown, column: number, criteria: Excel.FilterCriteria) => {
      filter.enabled = true;
      filter.criteria[column] = criteria;
    }),
    getRange: () => source,
  };
  const rowFields: string[] = [];
  const measures: string[] = [];
  const pivot = {
    id: 'minted-pivot-id',
    name: 'MintedPivot',
    load() {},
    hierarchies: { getItem: (field: string) => field },
    rowHierarchies: { add: (field: string) => rowFields.push(field) },
    dataHierarchies: {
      add: (field: string) => {
        measures.push(field);
        return { summarizeBy: '' };
      },
    },
    layout: {
      layoutType: '',
      subtotalLocation: '',
      showColumnGrandTotals: true,
      showRowGrandTotals: true,
      getRange: () => output,
      getDataBodyRange: () => data,
      getRowLabelRange: () => labels,
    },
  };
  const addPivot = vi.fn(() => pivot);
  let activeId = 'source-sheet';
  function sheet(id: string, name: string) {
    const item = {
      id,
      name,
      visibility: 'Visible',
      load() {},
      protection: {
        protected: false,
        load() {},
        protect: vi.fn(() => {
          item.protection.protected = true;
        }),
      },
      autoFilter: filter,
      pivotTables: { add: addPivot },
      getRange: (address: string) => (address === source.address ? source : target),
      activate: vi.fn(() => {
        activeId = id;
      }),
    };
    return item;
  }
  const sales = sheet('source-sheet', 'Sales');
  const report = sheet('target-sheet', 'Report');
  const sheets = {
    items: [sales, report],
    load() {},
    getItem: (name: string) => {
      const value = sheets.items.find((item) => item.name === name);
      if (!value) throw new Error('not found');
      return value;
    },
    getItemOrNullObject: (name: string) => {
      const value = sheets.items.find((item) => item.name === name);
      return value ?? { isNullObject: true, load() {} };
    },
    getActiveWorksheet: () => sheets.items.find((item) => item.id === activeId)!,
    add: vi.fn((name: string) => {
      const item = sheet('minted-sheet-id', name);
      sheets.items.push(item);
      return item;
    }),
  };
  const context = {
    workbook: { worksheets: sheets },
    sync: async () => {
      state.syncs++;
      if (state.failSync === state.syncs) throw new Error('simulated host connection lost');
    },
  };
  state.context = context;
  return {
    source,
    target,
    output,
    data,
    labels,
    filter,
    pivot,
    addPivot,
    rowFields,
    measures,
    sales,
    report,
    sheets,
  };
}
beforeEach(() => {
  state.syncs = 0;
  state.failSync = 0;
  vi.stubGlobal('Office', {
    DocumentMode: { ReadWrite: 'write' },
    context: {
      document: { url: 'https://example.test/book.xlsx', mode: 'write' },
      requirements: { isSetSupported: () => true },
    },
  });
});
afterEach(() => {
  vi.unstubAllGlobals();
});

describe('bounded native rectangle admission', () => {
  it('accepts explicit quoted finite ranges and refuses aliases, external books and oversized footprints', () => {
    expect(nativeRange("' Sales Q1 '!$B$2:$D$5")).toMatchObject({
      sheet: ' Sales Q1 ',
      rows: 4,
      columns: 3,
    });
    for (const selector of [
      'A1:B2',
      'Sales!A:A',
      'Sales!A1:AAA100',
      'Sales!B2:A1',
      '[Other]Sales!A1:B2',
      'Sales:Other!A1',
      'Sales!XFE1',
      'Sales!A1048577',
    ])
      expect(nativeRange(selector)).toBeUndefined();
  });
});

describe('native exact-value filters', () => {
  it('executes through the actual bridge dispatcher and verifies a literal wildcard value', async () => {
    const h = host();
    const result = await new ExcelBridge().actuate(
      request('filter-range', {
        filterRange: {
          range: h.source.address,
          column: 'Region',
          criterion1: '*East*',
          expectedSourceHash: await hash(h.source),
        },
      }),
    );
    expect(h.filter.apply).toHaveBeenCalledWith(h.source, 0, {
      filterOn: 'Values',
      values: ['*East*'],
    });
    expect(result).toMatchObject({
      ok: true,
      verification: { status: 'verified' },
      provenanceMissing: true,
      inverse: { op: 'not-reversible' },
    });
  });
  it('refuses stale evidence before mutation', async () => {
    const h = host();
    const captured = await hash(h.source);
    h.source.values[1]![1] = 99;
    expect(
      await applyNativeFilter(
        request('filter-range', {
          filterRange: {
            range: h.source.address,
            column: 'Region',
            criterion1: 'East',
            expectedSourceHash: captured,
          },
        }),
      ),
    ).toMatchObject({ ok: false, error: { code: 'source_conflict' } });
    expect(h.filter.apply).not.toHaveBeenCalled();
  });
  it('preserves existing filters and refuses ambiguous columns', async () => {
    const h = host();
    h.filter.enabled = true;
    const p = {
      range: h.source.address,
      column: 'Region',
      criterion1: 'East',
      expectedSourceHash: await hash(h.source),
    };
    expect(await applyNativeFilter(request('filter-range', { filterRange: p }))).toMatchObject({
      error: { code: 'filter_conflict' },
    });
    h.filter.enabled = false;
    h.source.values[0]![1] = 'Region';
    p.expectedSourceHash = await hash(h.source);
    expect(await applyNativeFilter(request('filter-range', { filterRange: p }))).toMatchObject({
      error: { code: 'unknown_column' },
    });
    expect(h.filter.apply).not.toHaveBeenCalled();
  });
  it('refuses compound criteria and unsupported hosts without any native write', async () => {
    const h = host();
    const p = {
      range: h.source.address,
      column: 'Region',
      criterion1: 'East',
      operator: 'Or',
      criterion2: 'West',
      expectedSourceHash: await hash(h.source),
    };
    expect(await applyNativeFilter(request('filter-range', { filterRange: p }))).toMatchObject({
      error: { code: 'unsupported_filter' },
    });
    vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => false } } });
    expect(
      await applyNativeFilter(
        request('filter-range', {
          filterRange: { ...p, operator: undefined, criterion2: undefined },
        }),
      ),
    ).toMatchObject({ error: { code: 'unsupported_host' } });
    expect(h.filter.apply).not.toHaveBeenCalled();
  });
  it('a mismatched native readback blocks another attempt', async () => {
    const h = host();
    h.filter.getRange = () => h.target;
    expect(
      await applyNativeFilter(
        request('filter-range', {
          filterRange: {
            range: h.source.address,
            column: 'Region',
            criterion1: 'East',
            expectedSourceHash: await hash(h.source),
          },
        }),
      ),
    ).toMatchObject({ ok: false, recoveryPending: true });
  });
});

describe('bounded native grouped pivots', () => {
  async function params(h: ReturnType<typeof host>) {
    return {
      sourceRange: h.source.address,
      destinationRange: h.target.address,
      expectedSourceHash: await hash(h.source),
      expectedDestinationHash: await hash(h.target),
      rowFields: ['Region'],
      valueFields: ['Revenue', 'Cost'],
      columnFields: [],
      filterFields: [],
    };
  }
  it('creates a native pivot, reads its totals, and records only the minted inverse identity', async () => {
    const h = host();
    const result = await new ExcelBridge().actuate(
      request('insert-pivot', { pivot: await params(h) }),
    );
    expect(h.addPivot).toHaveBeenCalledOnce();
    expect(h.rowFields).toEqual(['Region']);
    expect(h.measures).toEqual(['Revenue', 'Cost']);
    expect(result).toMatchObject({
      ok: true,
      location: 'Report!A1:C3',
      verification: { status: 'verified' },
      inverse: { op: 'delete-object', objectType: 'pivot', name: 'minted-pivot-id' },
    });
  });
  it('refuses occupied or changed destinations before creating an object', async () => {
    const h = host();
    const p = await params(h);
    h.target.values[0]![0] = 'manual';
    expect(await applyNativePivot(request('insert-pivot', { pivot: p }))).toMatchObject({
      error: { code: 'source_conflict' },
    });
    p.expectedDestinationHash = await hash(h.target);
    expect(await applyNativePivot(request('insert-pivot', { pivot: p }))).toMatchObject({
      error: { code: 'destination_collision' },
    });
    expect(h.addPivot).not.toHaveBeenCalled();
  });
  it('refuses unknown shapes and nonnumeric measures before host mutation', async () => {
    const h = host();
    const p = await params(h);
    expect(
      await applyNativePivot(
        request('insert-pivot', { pivot: { ...p, columnFields: ['Region'] } }),
      ),
    ).toMatchObject({ error: { code: 'unsupported_pivot' } });
    h.source.values[1]![1] = '9007199254740993';
    p.expectedSourceHash = await hash(h.source);
    expect(await applyNativePivot(request('insert-pivot', { pivot: p }))).toMatchObject({
      error: { code: 'unsupported_values' },
    });
    expect(h.addPivot).not.toHaveBeenCalled();
  });
  it('requires an explicit large enough reserved footprint and refuses same-sheet output', async () => {
    const h = host();
    const p = await params(h);
    expect(
      await applyNativePivot(
        request('insert-pivot', { pivot: { ...p, destinationRange: 'Report!A1:B2' } }),
      ),
    ).toMatchObject({ error: { code: 'invalid_footprint' } });
    expect(
      await applyNativePivot(
        request('insert-pivot', { pivot: { ...p, destinationRange: 'Sales!A1:D7' } }),
      ),
    ).toMatchObject({ error: { code: 'invalid_footprint' } });
    expect(h.addPivot).not.toHaveBeenCalled();
  });
  it('preserves the minted inverse after a configuration sync failure', async () => {
    const h = host();
    // Sync 1 resolves the two worksheets, sync 2 mints the pivot, sync 3 is its configuration.
    state.failSync = 3;
    expect(
      await applyNativePivot(request('insert-pivot', { pivot: await params(h) })),
    ).toMatchObject({ ok: false, recoveryPending: true, inverse: { name: 'minted-pivot-id' } });
    expect(h.addPivot).toHaveBeenCalledOnce();
  });
  it('does not report complete when totals or the landed footprint disagree', async () => {
    const h = host();
    h.data.values[0]![0] = 16;
    expect(
      await applyNativePivot(request('insert-pivot', { pivot: await params(h) })),
    ).toMatchObject({ ok: false, recoveryPending: true });
    h.data.values[0]![0] = 15;
    h.output.address = 'Report!A1:F20';
    expect(
      await applyNativePivot(request('insert-pivot', { pivot: await params(h) })),
    ).toMatchObject({ ok: false, recoveryPending: true });
  });
});

describe('native literal-list data validation', () => {
  it('uses the actual dispatcher, installs a native dropdown, and reads back its stopping rule', async () => {
    const h = host();
    const result = await new ExcelBridge().actuate(
      request('set-data-validation', {
        validation: {
          range: h.source.address,
          expectedSourceHash: await hash(h.source),
          values: ['Open', 'Closed'],
        },
      }),
    );
    expect(result).toMatchObject({
      ok: true,
      verification: { status: 'verified' },
      inverse: { op: 'not-reversible' },
    });
    expect(h.source.dataValidation.rule).toEqual({
      list: { source: 'Open,Closed', inCellDropDown: true },
    });
    expect(h.source.dataValidation.errorAlert).toMatchObject({ showAlert: true, style: 'Stop' });
  });
  it('refuses source drift, protected sheets, and existing validation before changing the rule', async () => {
    const h = host();
    const p = {
      range: h.source.address,
      expectedSourceHash: await hash(h.source),
      values: ['Open', 'Closed'],
    };
    h.source.values[1]![0] = 'changed';
    expect(
      await applyNativeValidation(request('set-data-validation', { validation: p })),
    ).toMatchObject({ error: { code: 'source_conflict' } });
    p.expectedSourceHash = await hash(h.source);
    h.sales.protection.protected = true;
    expect(
      await applyNativeValidation(request('set-data-validation', { validation: p })),
    ).toMatchObject({ error: { code: 'protected' } });
    h.sales.protection.protected = false;
    h.source.dataValidation.type = 'List';
    expect(
      await applyNativeValidation(request('set-data-validation', { validation: p })),
    ).toMatchObject({ error: { code: 'validation_conflict' } });
    expect(h.source.dataValidation.rule).toEqual({});
  });
  it('refuses formula sources, separators, duplicate values and oversized lists', async () => {
    const h = host();
    for (const values of [['=External!A1:A5'], ['Yes,No'], ['a', 'a'], ['x'.repeat(256)]])
      expect(
        await applyNativeValidation(
          request('set-data-validation', {
            validation: {
              range: h.source.address,
              expectedSourceHash: await hash(h.source),
              values,
            },
          }),
        ),
      ).toMatchObject({ ok: false, error: { code: 'invalid_validation' } });
    expect(h.source.dataValidation.rule).toEqual({});
  });
  it('an unconfirmed readback retains an uncertain outcome rather than reporting success', async () => {
    const h = host();
    Object.defineProperty(h.source.dataValidation, 'type', {
      get: () => (h.source.dataValidation.rule.list ? 'Custom' : 'None'),
      set() {},
    });
    expect(
      await applyNativeValidation(
        request('set-data-validation', {
          validation: {
            range: h.source.address,
            expectedSourceHash: await hash(h.source),
            values: ['Open', 'Closed'],
          },
        }),
      ),
    ).toMatchObject({ ok: false, recoveryPending: true });
  });
});

describe('native worksheet lifecycle identity', () => {
  it('creates at the end with readback and no promise of deletion undo', async () => {
    const h = host();
    expect(
      await new ExcelBridge().actuate(
        request('manage-worksheet', { worksheet: { action: 'create', name: 'Analysis' } }),
      ),
    ).toMatchObject({
      ok: true,
      verification: { status: 'verified' },
      inverse: { op: 'not-reversible' },
    });
    expect(h.sheets.add).toHaveBeenCalledWith('Analysis');
  });
  it('refuses replacement sheets with the same name but a different native identity', async () => {
    const h = host();
    expect(
      await applyNativeWorksheet(
        request('manage-worksheet', {
          worksheet: {
            action: 'rename',
            name: 'Sales',
            newName: 'Renamed',
            expectedId: 'old-sheet-id',
          },
        }),
      ),
    ).toMatchObject({ error: { code: 'target_conflict' } });
    expect(h.sales.name).toBe('Sales');
  });
  it('renames, activates and protects only the exact captured worksheet', async () => {
    const h = host();
    expect(
      await applyNativeWorksheet(
        request('manage-worksheet', {
          worksheet: {
            action: 'rename',
            name: 'Sales',
            newName: 'Renamed',
            expectedId: 'source-sheet',
          },
        }),
      ),
    ).toMatchObject({ ok: true });
    expect(h.sales.name).toBe('Renamed');
    expect(
      await applyNativeWorksheet(
        request('manage-worksheet', {
          worksheet: { action: 'activate', name: 'Report', expectedId: 'target-sheet' },
        }),
      ),
    ).toMatchObject({ ok: true });
    expect(h.report.activate).toHaveBeenCalledOnce();
    expect(
      await applyNativeWorksheet(
        request('manage-worksheet', {
          worksheet: { action: 'protect', name: 'Report', expectedId: 'target-sheet' },
        }),
      ),
    ).toMatchObject({ ok: true });
    expect(h.report.protection.protected).toBe(true);
  });
  it('refuses destructive delete, unprotect, duplicate names and missing identities', async () => {
    const h = host();
    for (const action of ['delete', 'unprotect'] as const)
      expect(
        await applyNativeWorksheet(
          request('manage-worksheet', {
            worksheet: { action, name: 'Sales', expectedId: 'source-sheet' },
          }),
        ),
      ).toMatchObject({ ok: false });
    expect(
      await applyNativeWorksheet(
        request('manage-worksheet', { worksheet: { action: 'create', name: 'sales' } }),
      ),
    ).toMatchObject({ error: { code: 'name_collision' } });
    expect(
      await applyNativeWorksheet(
        request('manage-worksheet', { worksheet: { action: 'activate', name: 'Sales' } }),
      ),
    ).toMatchObject({ error: { code: 'missing_identity' } });
    expect(h.sheets.add).not.toHaveBeenCalled();
  });
});
