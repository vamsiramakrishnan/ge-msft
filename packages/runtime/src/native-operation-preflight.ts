import {
  boundedExcelRange,
  CellSnapshotSchema,
  type ActuationRequest,
  type CellSnapshot,
} from '@ge/contracts';
import type { DocBridge } from './bridge.js';

/** Capture read-only native pins BEFORE approval. A supplied pin can be checked, never rebased. */
export async function prepareNativeOperationPreconditions(
  input: ActuationRequest,
  bridge: Pick<
    DocBridge,
    'surface' | 'captureCells' | 'captureCellsBatch' | 'captureWorksheetIdentity'
  >,
): Promise<ActuationRequest> {
  if (input.surface !== 'excel' || bridge.surface !== 'excel') return input;
  const request = structuredClone(input);
  const p = request.params;
  if (request.kind === 'manage-worksheet' && p.worksheet && p.worksheet.action !== 'create') {
    if (!bridge.captureWorksheetIdentity)
      throw new Error('This host cannot capture a worksheet identity before approval.');
    const target = await bridge.captureWorksheetIdentity(p.worksheet.name);
    if (
      !target.id ||
      target.id.length > 256 ||
      target.name !== p.worksheet.name ||
      (p.worksheet.expectedId && p.worksheet.expectedId !== target.id)
    )
      throw new Error('The worksheet identity changed before approval.');
    p.worksheet.expectedId = target.id;
    return request;
  }
  const table = request.kind === 'create-table' && p.table;
  const fields = table
    ? [{ selector: table.range, expected: table.expectedSourceHash }]
    : request.kind === 'insert-pivot' && p.pivot
      ? [
          { selector: p.pivot.sourceRange, expected: p.pivot.expectedSourceHash },
          { selector: p.pivot.destinationRange, expected: p.pivot.expectedDestinationHash },
        ]
      : request.kind === 'filter-range' && p.filterRange
        ? [{ selector: p.filterRange.range, expected: p.filterRange.expectedSourceHash }]
        : request.kind === 'set-data-validation' && p.validation
          ? [{ selector: p.validation.range, expected: p.validation.expectedSourceHash }]
          : [];
  if (!fields.length) return input;
  if (!bridge.captureCells)
    throw new Error('This host cannot capture native targets before approval.');
  const tableUsesActiveSheet = !!table && !table.range.includes('!');
  const bounds = fields.map((field) =>
    boundedExcelRange(tableUsesActiveSheet ? `Active!${field.selector}` : field.selector),
  );
  if (bounds.some((bound) => !bound))
    throw new Error(
      'Native operations require explicit sheet-qualified rectangles of at most 10,000 cells.',
    );
  const snapshots = bridge.captureCellsBatch
    ? await bridge.captureCellsBatch(fields.map((field) => field.selector))
    : await Promise.all(fields.map((field) => bridge.captureCells!(field.selector)));
  if (snapshots.length !== fields.length)
    throw new Error('Native target capture did not return every requested rectangle.');
  const checked: CellSnapshot[] = snapshots.map((snapshot, index) => {
    const parsed = CellSnapshotSchema.parse(snapshot);
    const actual = boundedExcelRange(parsed.locator);
    const expected = bounds[index]!;
    if (
      parsed.surface !== 'excel' ||
      !parsed.objectId ||
      !actual ||
      (!tableUsesActiveSheet && actual.sheet !== expected.sheet) ||
      actual.firstRow !== expected.firstRow ||
      actual.firstColumn !== expected.firstColumn ||
      actual.rows !== expected.rows ||
      actual.columns !== expected.columns ||
      parsed.values.length !== expected.rows ||
      parsed.values[0]!.length !== expected.columns ||
      new TextEncoder().encode(JSON.stringify(parsed)).byteLength > 1024 * 1024
    )
      throw new Error('Native capture returned a different or oversized target.');
    if (fields[index]!.expected && fields[index]!.expected !== parsed.hash)
      throw new Error('The captured native target changed. Capture it again before approval.');
    return parsed;
  });
  if (checked.some((snapshot) => snapshot.documentId !== checked[0]!.documentId))
    throw new Error('Native targets belong to different documents.');
  if (table) {
    table.range = checked[0]!.locator;
    table.expectedSourceHash = checked[0]!.hash;
  } else if (request.kind === 'insert-pivot' && p.pivot) {
    p.pivot.expectedSourceHash = checked[0]!.hash;
    p.pivot.expectedDestinationHash = checked[1]!.hash;
  } else if (request.kind === 'filter-range' && p.filterRange)
    p.filterRange.expectedSourceHash = checked[0]!.hash;
  else if (request.kind === 'set-data-validation' && p.validation)
    p.validation.expectedSourceHash = checked[0]!.hash;
  return request;
}
