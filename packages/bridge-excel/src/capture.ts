import type { ResolvedContext } from '@ge/contracts';
import {
  native,
  toContextNative,
  type Block,
  type NativeContent,
  type ToContextOptions,
} from '@ge/content';

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

/**
 * A used range → a single native table `Block` for the `<doc_state>` snapshot (ADR-0003). Same
 * header/data split as `rangeToContext`, anchored on a `range:<address>` locator so the snapshot
 * inventory carries a stable id. Empty grid → `[]`.
 */
export function usedRangeToBlocks(address: string, values: string[][]): Block[] {
  const { columns, rows } = splitHeaderRows(values);
  if (columns.length === 0) return [];
  return [native.table({ columns, rows }, `range:${address}`)];
}

/** Cap lazy `search_document` row reads so a common term can't blow the per-turn budget. */
export const MAX_SEARCH_ROWS = 8;

/**
 * Scan a used range's grid for rows containing `query` (case-insensitive substring on any cell),
 * and return the matching rows — with the header row preserved — as content via `rangeToContext`.
 * Bounded to the top {@link MAX_SEARCH_ROWS} matches. Empty query / no header / no match → `[]`.
 * Pure: the host read happens in the bridge; this is the match + shaping step.
 */
export function searchUsedRange(
  address: string,
  values: string[][],
  query: string,
): ResolvedContext[] {
  const needle = query.trim().toLowerCase();
  if (!needle) return [];
  const header = values[0];
  if (!header) return [];

  const matched: string[][] = [];
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    if (!row) continue;
    if (
      row.some((cell) =>
        String(cell ?? '')
          .toLowerCase()
          .includes(needle),
      )
    ) {
      matched.push(row);
      if (matched.length >= MAX_SEARCH_ROWS) break;
    }
  }
  if (matched.length === 0) return [];
  return rangeToContext(address, [header, ...matched]);
}
