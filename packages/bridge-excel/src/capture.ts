import type { ResolvedContext } from '@ge/contracts';
import { native, toContextNative, type NativeContent, type ToContextOptions } from '@ge/content';

/**
 * Pure mapping from an Excel range's address + 2D values into grounding-ready context — no
 * Office.js here, so it's unit-testable. The `ExcelBridge` reads a range's `.address` and
 * `.values` via `Excel.run` and hands the raw grid to these functions; they go straight
 * through `@ge/content` (native path, no Markdown round-trip) as a table block and carry a
 * `range:<address>` write-back locator.
 */

/** Split a 2D grid into a header row (row 0) and the remaining data rows. */
export function splitHeaderRows(values: string[][]): { columns: string[]; rows: string[][] } {
  const header = values[0];
  if (!header) return { columns: [], rows: [] };
  return { columns: header, rows: values.slice(1) };
}

/**
 * A range → a single native table block, anchored to its address, then chunked. Row 0 is
 * treated as the header; everything below is data.
 */
export function rangeToContext(
  address: string,
  values: string[][],
  opts: ToContextOptions = {},
): ResolvedContext[] {
  const { columns, rows } = splitHeaderRows(values);
  if (columns.length === 0) return [];
  const content: NativeContent = {
    sourceId: `xl:${address}`,
    surface: 'excel',
    title: address,
    blocks: [native.table({ columns, rows }, `range:${address}`)],
  };
  return toContextNative(content, opts);
}

/** The current selection's grid → context (same table mapping as `rangeToContext`). */
export function selectionValuesToContext(address: string, values: string[][]): ResolvedContext[] {
  return rangeToContext(address, values);
}
