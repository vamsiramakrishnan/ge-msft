import type { ActuationRequest, ActuationResult, CellSnapshot } from '@ge/contracts';
import { excelDocumentId, snapshotRange } from './cell-snapshot.js';
import { isSet } from './capabilities-runtime.js';
import { nativeRun } from './native-run.js';
import { resolveWorksheet, resolveWorksheetPair } from './worksheet-lookup.js';

function hasControl(value: string): boolean {
  return [...value].some((character) => character.charCodeAt(0) < 32);
}
function invalidSheetCharacters(value: string): boolean {
  return (
    hasControl(value) ||
    ['[', ']', ':', '*', '?', '/', '\\'].some((character) => value.includes(character))
  );
}

export { boundedExcelRange as nativeRange } from '@ge/contracts';
import { boundedExcelRange as nativeRange } from '@ge/contracts';

function refusal(req: ActuationRequest, code: string, message: string): ActuationResult {
  return { ok: false, changeId: req.changeId, kind: req.kind, error: { code, message } };
}
function uncertain(req: ActuationRequest, inverse?: ActuationResult['inverse']): ActuationResult {
  return {
    ...refusal(
      req,
      'verification_failed',
      'The native outcome could not be verified. Inspect the document before another attempt.',
    ),
    recoveryPending: true,
    ...(inverse ? { inverse } : {}),
  };
}
function byteSize(value: unknown): number {
  return new TextEncoder().encode(JSON.stringify(value)).byteLength;
}
function blank(snapshot: CellSnapshot): boolean {
  return (
    snapshot.values.every((row) => row.every((value) => value === '' || value === null)) &&
    !snapshot.formulas?.some((row) => row.some(Boolean))
  );
}
function boundedSnapshot(snapshot: CellSnapshot): boolean {
  return byteSize(snapshot) <= 1024 * 1024;
}

/** Install literal choices only; no formula is accepted as a validation source. */
export async function applyNativeValidation(req: ActuationRequest): Promise<ActuationResult> {
  const p = req.params.validation;
  const bounds = p && nativeRange(p.range);
  if (!p || !bounds || !p.expectedSourceHash)
    return refusal(
      req,
      'invalid_target',
      'Validation needs an explicit bounded rectangle and expectedSourceHash from captureCells.',
    );
  if (
    !p.values.length ||
    p.values.length > 32 ||
    p.values.some(
      (value) =>
        !value ||
        value.length > 64 ||
        [',', ';', '"'].some((separator) => value.includes(separator)) ||
        hasControl(value) ||
        /^[\s]*[=+\-@]/.test(value),
    ) ||
    new Set(p.values).size !== p.values.length ||
    p.values.join(',').length > 255
  )
    return refusal(
      req,
      'invalid_validation',
      'Use unique literal list items without formulas or separators, at most 255 characters in total.',
    );
  if (!isSet('ExcelApi', '1.8'))
    return refusal(req, 'unsupported_host', 'Native list validation requires ExcelApi 1.8.');
  const documentId = await excelDocumentId();
  return nativeRun({ undoGroup: true }, async (ctx) => {
    const sheet = await resolveWorksheet(ctx, bounds.sheet);
    const range = sheet.getRange(p.range);
    const validation = range.dataValidation;
    sheet.protection.load('protected');
    validation.load('type');
    const source = await snapshotRange(range, ctx, documentId);
    if (!boundedSnapshot(source))
      return refusal(req, 'capture_limit', 'The target exceeds the native-operation byte budget.');
    if (source.hash !== p.expectedSourceHash)
      return refusal(
        req,
        'source_conflict',
        'The captured target changed before validation could be installed.',
      );
    if (sheet.protection.protected) return refusal(req, 'protected', 'The sheet is protected.');
    if (validation.type !== 'None')
      return refusal(
        req,
        'validation_conflict',
        'Existing validation rules are preserved; replacing them requires another reviewed operation.',
      );
    const choices = p.values.join(',');
    validation.rule = { list: { inCellDropDown: true, source: choices } };
    validation.errorAlert = {
      showAlert: true,
      style: 'Stop',
      title: 'Choose a listed value',
      message: 'Select a value from the approved list.',
    };
    validation.load('type,rule,errorAlert');
    await ctx.sync();
    const inverse = {
      op: 'not-reversible' as const,
      reason: 'The applied validation rule is recorded; automatic restoration is unavailable.',
    };
    if (
      String(validation.type) !== 'List' ||
      validation.rule.list?.source !== choices ||
      validation.rule.list.inCellDropDown !== true ||
      validation.errorAlert.showAlert !== true ||
      validation.errorAlert.style !== 'Stop'
    )
      return uncertain(req, inverse);
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: source.locator,
      inverse,
      verification: { status: 'verified' },
    };
  });
}

/** Native AutoFilter only. Existing filters are preserved by refusing replacement. */
export async function applyNativeFilter(req: ActuationRequest): Promise<ActuationResult> {
  const p = req.params.filterRange;
  const bounds = p && nativeRange(p.range);
  if (!p || !bounds || bounds.rows < 2 || !p.expectedSourceHash)
    return refusal(
      req,
      'invalid_target',
      'Filter needs a sheet-qualified bounded rectangle and expectedSourceHash from captureCells.',
    );
  if (p.operator || p.criterion2 || p.criterion1.length > 512 || !p.column || p.column.length > 256)
    return refusal(
      req,
      'unsupported_filter',
      'This filter supports one exact text value. Compound and wildcard criteria require another operation.',
    );
  if (!isSet('ExcelApi', '1.9'))
    return refusal(req, 'unsupported_host', 'Native range filters require ExcelApi 1.9.');
  const documentId = await excelDocumentId();
  return nativeRun({ undoGroup: true }, async (ctx) => {
    const sheet = await resolveWorksheet(ctx, bounds.sheet);
    const range = sheet.getRange(p.range);
    const filter = sheet.autoFilter;
    filter.load('enabled');
    sheet.protection.load('protected');
    const source = await snapshotRange(range, ctx, documentId);
    if (!boundedSnapshot(source))
      return refusal(req, 'capture_limit', 'The source exceeds the native-operation byte budget.');
    if (source.hash !== p.expectedSourceHash)
      return refusal(
        req,
        'source_conflict',
        'The reviewed source changed. Capture it again before approval.',
      );
    if (sheet.protection.protected) return refusal(req, 'protected', 'The sheet is protected.');
    if (filter.enabled)
      return refusal(
        req,
        'filter_conflict',
        'An existing filter must be reviewed and cleared explicitly before adding another.',
      );
    const headers = source.values[0]!.map(String);
    const matches = headers.flatMap((header, index) => (header === p.column ? [index] : []));
    if (matches.length !== 1)
      return refusal(req, 'unknown_column', 'The exact header must identify one column.');
    const column = matches[0]!;
    // Values filtering treats wildcard characters as data, unlike a custom criterion string.
    filter.apply(range, column, { filterOn: 'Values', values: [p.criterion1] });
    filter.load('enabled,criteria');
    const landed = filter.getRange();
    landed.load('address');
    await ctx.sync();
    const criteria = filter.criteria[column];
    const expected = nativeRange(landed.address);
    const sameRange =
      expected &&
      expected.sheet === bounds.sheet &&
      expected.firstRow === bounds.firstRow &&
      expected.firstColumn === bounds.firstColumn &&
      expected.rows === bounds.rows &&
      expected.columns === bounds.columns;
    const inverse = {
      op: 'not-reversible' as const,
      reason: 'The applied filter is recorded; automatic filter restoration is unavailable.',
    };
    if (
      !filter.enabled ||
      !sameRange ||
      criteria?.filterOn !== 'Values' ||
      criteria.values?.length !== 1 ||
      criteria.values[0] !== p.criterion1
    )
      return uncertain(req, inverse);
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: landed.address,
      inverse,
      verification: { status: 'verified' },
    };
  });
}

/** A bounded native grouped SUM; no substitute cell grid is ever generated. */
export async function applyNativePivot(req: ActuationRequest): Promise<ActuationResult> {
  const p = req.params.pivot;
  const sourceBounds = p && nativeRange(p.sourceRange);
  const destination = p && nativeRange(p.destinationRange);
  if (
    !p ||
    !sourceBounds ||
    !destination ||
    sourceBounds.rows < 2 ||
    !p.expectedSourceHash ||
    !p.expectedDestinationHash
  )
    return refusal(
      req,
      'invalid_target',
      'Pivot needs explicit source/destination rectangles and both captured snapshot hashes.',
    );
  if (
    p.rowFields.length !== 1 ||
    p.valueFields.length < 1 ||
    p.valueFields.length > 4 ||
    p.columnFields.length ||
    p.filterFields.length ||
    p.name
  )
    return refusal(
      req,
      'unsupported_pivot',
      'Use one row field and one to four numeric SUM fields. Column/filter fields and custom object names are not supported.',
    );
  if (
    sourceBounds.sheet === destination.sheet ||
    destination.rows < sourceBounds.rows + 3 ||
    destination.columns < p.valueFields.length + 2
  )
    return refusal(
      req,
      'invalid_footprint',
      'Reserve an empty rectangle on a different worksheet with at least source rows + 3 rows and measure count + 2 columns.',
    );
  if (!isSet('ExcelApi', '1.8'))
    return refusal(req, 'unsupported_host', 'Native PivotTables require ExcelApi 1.8.');
  const documentId = await excelDocumentId();
  return nativeRun({ undoGroup: true }, async (ctx) => {
    const [sourceSheet, sheet] = await resolveWorksheetPair(
      ctx,
      sourceBounds.sheet,
      destination.sheet,
    );
    const sourceRange = sourceSheet.getRange(p.sourceRange);
    const destinationRange = sheet.getRange(p.destinationRange);
    sheet.protection.load('protected');
    const source = await snapshotRange(sourceRange, ctx, documentId);
    const target = await snapshotRange(destinationRange, ctx, documentId);
    if (!boundedSnapshot(source) || !boundedSnapshot(target))
      return refusal(req, 'capture_limit', 'Sources exceed the native-operation byte budget.');
    if (source.hash !== p.expectedSourceHash || target.hash !== p.expectedDestinationHash)
      return refusal(req, 'source_conflict', 'Source or destination changed after capture.');
    if (!blank(target))
      return refusal(
        req,
        'destination_collision',
        'The complete reserved destination must be empty.',
      );
    if (sheet.protection.protected)
      return refusal(req, 'protected', 'The destination sheet is protected.');
    const headers = source.values[0]!.map(String);
    if (
      headers.some((header) => !header || header.length > 256) ||
      new Set(headers).size !== headers.length
    )
      return refusal(req, 'invalid_headers', 'Pivot headers must be nonempty, unique and bounded.');
    const rowIndex = headers.indexOf(p.rowFields[0]!);
    const valueIndexes = p.valueFields.map((field) => headers.indexOf(field));
    if (
      rowIndex < 0 ||
      valueIndexes.some((index) => index < 0 || index === rowIndex) ||
      new Set(valueIndexes).size !== valueIndexes.length
    )
      return refusal(
        req,
        'unknown_field',
        'Choose distinct row and value fields from exact source headers.',
      );
    const expected = new Map<string, number[]>();
    for (const row of source.values.slice(1)) {
      const key = row[rowIndex];
      const amounts = valueIndexes.map((index) => row[index]);
      if (
        typeof key !== 'string' ||
        !key ||
        key.length > 256 ||
        amounts.some((value) => typeof value !== 'number' || !Number.isFinite(value))
      )
        return refusal(
          req,
          'unsupported_values',
          'Native grouped SUM requires nonempty text row labels and finite numeric measures.',
        );
      const sums = expected.get(key) ?? amounts.map(() => 0);
      amounts.forEach((value, index) => (sums[index] = sums[index]! + (value as number)));
      if (sums.some((value) => !Number.isFinite(value)))
        return refusal(req, 'invalid_totals', 'The grouped totals are not finite.');
      expected.set(key, sums);
    }
    // The chosen layout has no subtotals, totals, filters or column dimensions. Its maximum
    // footprint is bounded by source rows + header padding, reserved explicitly above.
    const anchor = destinationRange.getCell(0, 0);
    const pivot = sheet.pivotTables.add(
      `GE_${crypto.randomUUID().replace(/-/g, '')}`,
      sourceRange,
      anchor,
    );
    pivot.load('id,name');
    await ctx.sync();
    const inverse = {
      op: 'delete-object' as const,
      objectType: 'pivot' as const,
      name: pivot.id,
    };
    try {
      pivot.rowHierarchies.add(pivot.hierarchies.getItem(p.rowFields[0]!));
      p.valueFields.forEach((field) => {
        pivot.dataHierarchies.add(pivot.hierarchies.getItem(field)).summarizeBy = 'Sum';
      });
      pivot.layout.layoutType = 'Tabular';
      pivot.layout.subtotalLocation = 'Off';
      pivot.layout.showColumnGrandTotals = false;
      pivot.layout.showRowGrandTotals = false;
      await ctx.sync();
      const output = pivot.layout.getRange();
      const data = pivot.layout.getDataBodyRange();
      const labels = pivot.layout.getRowLabelRange();
      output.load('address,rowCount,columnCount');
      data.load('values,rowCount,columnCount');
      labels.load('values,rowCount,columnCount');
      await ctx.sync();
      const landed = nativeRange(output.address);
      const inside =
        landed &&
        landed.sheet === destination.sheet &&
        landed.firstRow >= destination.firstRow &&
        landed.firstColumn >= destination.firstColumn &&
        landed.firstRow + landed.rows <= destination.firstRow + destination.rows &&
        landed.firstColumn + landed.columns <= destination.firstColumn + destination.columns;
      // RowLabelRange may include its field header; align its last data.rowCount labels.
      const rowLabels = labels.values.slice(-data.rowCount).map((row) => String(row[0] ?? ''));
      const seen = new Set<string>();
      const totalsMatch =
        data.rowCount === expected.size &&
        data.columnCount === valueIndexes.length &&
        data.values.every((row, index) => {
          const key = rowLabels[index]!;
          const sums = expected.get(key);
          if (!sums || seen.has(key)) return false;
          seen.add(key);
          return (
            row.length === sums.length &&
            row.every(
              (value, column) =>
                typeof value === 'number' &&
                Math.abs(value - sums[column]!) <= Math.max(1, Math.abs(sums[column]!)) * 1e-10,
            )
          );
        });
      if (!inside || !totalsMatch) return uncertain(req, inverse);
      return {
        ok: true,
        changeId: req.changeId,
        kind: req.kind,
        location: output.address,
        inverse,
        verification: { status: 'verified' },
      };
    } catch {
      // The minted object may already exist. Preserve its inverse; never retry or claim rollback.
      return uncertain(req, inverse);
    }
  });
}

/** Sheet IDs are required for existing-sheet actions so a renamed replacement is never targeted. */
export async function applyNativeWorksheet(req: ActuationRequest): Promise<ActuationResult> {
  const p = req.params.worksheet;
  const validName = (name: string | undefined): name is string =>
    !!name?.trim() &&
    name.length <= 31 &&
    !invalidSheetCharacters(name) &&
    !name.startsWith("'") &&
    !name.endsWith("'");
  if (!p || !validName(p.name) || (p.position && p.position !== 'end'))
    return refusal(
      req,
      'invalid_target',
      'Provide a valid exact worksheet name. Only end placement is supported.',
    );
  if (p.action === 'rename' && !validName(p.newName))
    return refusal(req, 'invalid_name', 'Rename requires a valid new worksheet name.');
  if (p.action !== 'create' && !p.expectedId)
    return refusal(
      req,
      'missing_identity',
      'Capture and supply the exact expectedId for this worksheet before approval.',
    );
  if (!isSet('ExcelApi', '1.2'))
    return refusal(req, 'unsupported_host', 'Worksheet lifecycle checks require ExcelApi 1.2.');
  return nativeRun({ undoGroup: true }, async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    // Collection metadata is bounded; content is never fetched until a finite empty-delete check.
    sheets.load('items/id,items/name,items/visibility');
    await ctx.sync();
    if (sheets.items.length > 256)
      return refusal(req, 'capture_limit', 'Worksheet inventory exceeds the lifecycle budget.');
    const existing = sheets.items.find(
      (sheet) => sheet.name.toLowerCase() === p.name.toLowerCase(),
    );
    if (p.action === 'create') {
      if (existing)
        return refusal(req, 'name_collision', 'A worksheet with that name already exists.');
      const created = sheets.add(p.name);
      created.load('id,name');
      await ctx.sync();
      return {
        ok: true,
        changeId: req.changeId,
        kind: req.kind,
        location: created.name,
        inverse: {
          op: 'not-reversible',
          reason: 'A created worksheet may acquire user data; automatic deletion is not offered.',
        },
        verification: { status: created.id && created.name === p.name ? 'verified' : 'unknown' },
      };
    }
    if (!existing || existing.id !== p.expectedId || existing.name !== p.name)
      return refusal(req, 'target_conflict', 'Worksheet identity or exact name changed.');
    if (
      p.action === 'rename' &&
      sheets.items.some(
        (sheet) =>
          sheet.id !== existing.id && sheet.name.toLowerCase() === p.newName!.toLowerCase(),
      )
    )
      return refusal(req, 'name_collision', 'The new name already exists.');
    if (p.action === 'delete')
      return refusal(
        req,
        'irreversible_delete',
        'Worksheet deletion is not exposed: dependencies and a complete recovery image cannot be guaranteed.',
      );
    existing.protection.load('protected');
    await ctx.sync();
    if (p.action === 'unprotect')
      return refusal(
        req,
        'unsupported_action',
        'Removing protection requires an explicit host workflow and is not automated.',
      );
    if (p.action === 'protect' && existing.protection.protected)
      return refusal(req, 'protection_conflict', 'The worksheet is already protected.');
    if (p.action === 'rename') existing.name = p.newName!;
    else if (p.action === 'activate') existing.activate();
    else if (p.action === 'protect') existing.protection.protect();
    existing.load('id,name');
    existing.protection.load('protected');
    const active = sheets.getActiveWorksheet();
    active.load('id');
    await ctx.sync();
    const verified =
      existing.id === p.expectedId &&
      (p.action === 'rename'
        ? existing.name === p.newName
        : p.action === 'activate'
          ? active.id === p.expectedId
          : existing.protection.protected);
    const inverse = {
      op: 'not-reversible' as const,
      reason:
        'Automatic lifecycle restoration is unavailable; inspect the recorded native worksheet identity.',
    };
    return verified
      ? {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: existing.name,
          inverse,
          verification: { status: 'verified' },
        }
      : uncertain(req, inverse);
  });
}
