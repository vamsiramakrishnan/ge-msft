import { describe, expect, it, vi } from 'vitest';
import { asChangeId, makeCellSnapshot, type ActuationRequest } from '@ge/contracts';
import { prepareNativeOperationPreconditions } from './native-operation-preflight.js';

function request(
  kind: ActuationRequest['kind'],
  params: ActuationRequest['params'],
): ActuationRequest {
  return { kind, surface: 'excel', changeId: asChangeId('preflight'), params };
}
async function snapshot(locator = 'Sales!A1:B2') {
  return makeCellSnapshot({
    surface: 'excel',
    documentId: 'book',
    objectId: 'native-sheet-id',
    locator,
    values: [
      ['Region', 'Revenue'],
      ['East', 10],
    ],
    formulas: [
      ['', ''],
      ['', ''],
    ],
  });
}
describe('native operation preapproval captures', () => {
  it('canonicalizes an active-sheet table source before review and pins its source', async () => {
    const captured = await snapshot();
    const output = await prepareNativeOperationPreconditions(
      request('create-table', { table: { range: 'A1:B2', hasHeaders: true } }),
      {
        surface: 'excel',
        captureCells: async (selector) => {
          expect(selector).toBe('A1:B2');
          return captured;
        },
      },
    );
    expect(output.params.table).toMatchObject({
      range: 'Sales!A1:B2',
      expectedSourceHash: captured.hash,
    });
  });

  it('adds a content pin to the exact filter target without mutating caller input or dispatching', async () => {
    const captured = await snapshot();
    const captureCells = vi.fn(async () => captured);
    const input = request('filter-range', {
      filterRange: { range: 'Sales!A1:B2', column: 'Region', criterion1: 'East' },
    });
    const result = await prepareNativeOperationPreconditions(input, {
      surface: 'excel',
      captureCells,
    });
    expect(captureCells).toHaveBeenCalledWith('Sales!A1:B2');
    expect(result.params.filterRange?.expectedSourceHash).toBe(captured.hash);
    expect(input.params.filterRange?.expectedSourceHash).toBeUndefined();
  });
  it('checks a supplied source pin instead of silently replacing a stale approved revision', async () => {
    const captured = await snapshot();
    const input = request('set-data-validation', {
      validation: {
        range: captured.locator,
        values: ['Open'],
        expectedSourceHash: `sha256:${'0'.repeat(64)}`,
      },
    });
    await expect(
      prepareNativeOperationPreconditions(input, {
        surface: 'excel',
        captureCells: async () => captured,
      }),
    ).rejects.toThrow('changed');
    expect(input.params.validation?.expectedSourceHash).toBe(`sha256:${'0'.repeat(64)}`);
  });
  it('captures pivot source and destination together and preserves their order and native identity', async () => {
    const source = await snapshot();
    const destination = await snapshot('Report!D4:E5');
    const captureCells = vi.fn();
    const captureCellsBatch = vi.fn(async () => [source, destination]);
    const input = request('insert-pivot', {
      pivot: {
        sourceRange: source.locator,
        destinationRange: destination.locator,
        rowFields: ['Region'],
        valueFields: ['Revenue'],
        columnFields: [],
        filterFields: [],
      },
    });
    const output = await prepareNativeOperationPreconditions(input, {
      surface: 'excel',
      captureCells,
      captureCellsBatch,
    });
    expect(captureCellsBatch).toHaveBeenCalledWith([source.locator, destination.locator]);
    expect(captureCells).not.toHaveBeenCalled();
    expect(output.params.pivot).toMatchObject({
      expectedSourceHash: source.hash,
      expectedDestinationHash: destination.hash,
    });
  });
  it('rejects capture widening, reordered ranges and cross-document results', async () => {
    const source = await snapshot();
    const destination = await snapshot('Report!D4:E5');
    const input = request('insert-pivot', {
      pivot: {
        sourceRange: source.locator,
        destinationRange: destination.locator,
        rowFields: ['Region'],
        valueFields: ['Revenue'],
        columnFields: [],
        filterFields: [],
      },
    });
    for (const result of [
      [destination, source],
      [source, { ...destination, documentId: 'another-book' }],
      [source],
    ]) {
      await expect(
        prepareNativeOperationPreconditions(input, {
          surface: 'excel',
          captureCells: async () => source,
          captureCellsBatch: async () => result,
        }),
      ).rejects.toThrow();
    }
  });
  it('does not capture unqualified, external, whole-column or overbudget selectors', async () => {
    const captureCells = vi.fn();
    for (const range of ['A1:B2', 'Sales!A:A', '[Foreign]Sales!A1:B2', 'Sales!A1:Z1000'])
      await expect(
        prepareNativeOperationPreconditions(
          request('filter-range', { filterRange: { range, column: 'Region', criterion1: 'East' } }),
          { surface: 'excel', captureCells },
        ),
      ).rejects.toThrow('explicit');
    expect(captureCells).not.toHaveBeenCalled();
  });
  it('pins existing worksheets using metadata without reading any cell content', async () => {
    const captureCells = vi.fn();
    const captureWorksheetIdentity = vi.fn(async (name) => ({ id: 'native-sheet-id', name }));
    const output = await prepareNativeOperationPreconditions(
      request('manage-worksheet', {
        worksheet: { action: 'rename', name: 'Sales', newName: 'Sales2' },
      }),
      { surface: 'excel', captureCells, captureWorksheetIdentity },
    );
    expect(output.params.worksheet?.expectedId).toBe('native-sheet-id');
    expect(captureWorksheetIdentity).toHaveBeenCalledWith('Sales');
    expect(captureCells).not.toHaveBeenCalled();
    await expect(
      prepareNativeOperationPreconditions(
        request('manage-worksheet', {
          worksheet: { action: 'activate', name: 'Sales', expectedId: 'old-id' },
        }),
        { surface: 'excel', captureWorksheetIdentity },
      ),
    ).rejects.toThrow('identity changed');
  });
  it('fails closed when identity capture is unavailable and never invents it from a display label', async () => {
    await expect(
      prepareNativeOperationPreconditions(
        request('manage-worksheet', { worksheet: { action: 'activate', name: 'Sales' } }),
        { surface: 'excel' },
      ),
    ).rejects.toThrow('cannot capture');
    const input = request('manage-worksheet', { worksheet: { action: 'create', name: 'New' } });
    expect(await prepareNativeOperationPreconditions(input, { surface: 'excel' })).toEqual(input);
  });
});
