import type { ContextRef, ResolvedContext } from '@ge/contracts';
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

/**
 * Split a 2D grid into a header row and the remaining data rows. Skips any leading rows that are
 * entirely blank (every cell empty/whitespace) before treating a row as the header — a blank
 * formatting row above the real header (common when `getUsedRange` picks up formatted-but-empty
 * rows, or a sheet has a visual gap row) would otherwise be captured as an all-empty header,
 * silently corrupting every downstream column name.
 */
export function splitHeaderRows(values: string[][]): { columns: string[]; rows: string[][] } {
  // Excel returns numbers and booleans in `values` despite the string typing; a numeric first row
  // made `.trim()` throw, so every read of a numbers-only range failed (live 2026-09-30).
  const text = (cell: unknown) => String(cell ?? '').trim();
  const headerIdx = values.findIndex((row) => row.some((cell) => text(cell) !== ''));
  if (headerIdx === -1) return { columns: [], rows: [] };
  return { columns: values[headerIdx]!, rows: values.slice(headerIdx + 1) };
}

/** A row whose every non-blank cell is a number: data, never a header. */
function isNumericRow(row: readonly unknown[]): boolean {
  const cells = row.map((cell) => String(cell ?? '').trim()).filter(Boolean);
  return cells.length > 0 && cells.every((cell) => Number.isFinite(Number(cell)));
}

function columnLetters(index: number): string {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/**
 * A headerless grid (its first row is numbers) as a table the model can address: its columns are
 * named by their sheet letters and each row carries its sheet row number. Reading `G2:G11` used to
 * turn `130000` into a column name and leave the model counting rows; it then commented the wrong
 * cell.
 */
function headerlessTable(
  address: string,
  values: string[][],
  rowNumbers?: readonly number[],
): { columns: string[]; rows: string[][] } {
  // The cell part follows the last `!`: a sheet name such as `FY2026Data` must not be read as a cell.
  const cell = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})/.exec(address.slice(address.lastIndexOf('!') + 1));
  const firstCol = cell
    ? cell[1]!
        .toUpperCase()
        .split('')
        .reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
    : 0;
  const firstRow = cell ? Number(cell[2]) : 1;
  const width = Math.max(...values.map((row) => row.length));
  return {
    columns: ['row', ...Array.from({ length: width }, (_, i) => columnLetters(firstCol + i))],
    rows: values.map((row, i) => [
      String(rowNumbers?.[i] ?? firstRow + i),
      ...row.map((v) => String(v ?? '')),
    ]),
  };
}

/**
 * A range → a single native table block, anchored to its address, then chunked. Row 0 is
 * treated as the header; everything below is data.
 */
export function rangeToContext(
  address: string,
  values: string[][],
  opts: ToContextOptions = {},
  /** Sheet row of each entry in `values`, when the rows are not contiguous (search matches). */
  rowNumbers?: readonly number[],
): ResolvedContext[] {
  const { columns, rows } =
    values.length > 0 && isNumericRow(values[0]!)
      ? headerlessTable(address, values, rowNumbers)
      : splitHeaderRows(values);
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
  const matchedAt: number[] = [];
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
      matchedAt.push(i);
      if (matched.length >= MAX_SEARCH_ROWS) break;
    }
  }
  if (matched.length === 0) return [];
  // Matched rows are not contiguous: label each with its real sheet row (security review).
  const start = /^\$?[A-Za-z]{1,3}\$?(\d{1,7})/.exec(address.slice(address.lastIndexOf('!') + 1));
  const firstRow = start ? Number(start[1]) : 1;
  return rangeToContext(address, [header, ...matched], {}, [
    firstRow,
    ...matchedAt.map((i) => firstRow + i),
  ]);
}

/** One reply in an Excel comment thread. */
export interface ExcelCommentReply {
  readonly authorName: string;
  readonly content: string;
}

/**
 * A comment thread read from the workbook: the comment, the cell it sits on and its replies. All
 * of it is untrusted workbook content — the runtime carries it as data.
 */
export interface ExcelComment {
  readonly id: string;
  readonly authorName: string;
  readonly content: string;
  /** `undefined` when the host can't report it (below ExcelApi 1.11). */
  readonly resolved: boolean | undefined;
  /** Sheet-qualified cell address, e.g. `Sales!B4` (empty if the host couldn't report it). */
  readonly address: string;
  readonly replies: readonly ExcelCommentReply[];
}

function commentAuthor(name: string): string {
  return name.trim() || 'Unknown author';
}

/**
 * Host text as a single JSON-quoted line: newlines and quotes are escaped, so a comment body can't
 * forge a builder-owned line (a fake thread header or "Reply from …") or break out of its quotes.
 */
function quoted(text: string): string {
  return JSON.stringify(text);
}

/** The host comment id an Excel `comment` ref points at (its `xl:comment:` id). */
export function commentIdFromRef(ref: ContextRef): string | undefined {
  const prefix = 'xl:comment:';
  return ref.id.startsWith(prefix) ? ref.id.slice(prefix.length) || undefined : undefined;
}

/**
 * Comment threads → attachable `comment` refs, one per thread. The id carries the host comment id
 * the `reply <commentId> "text"` verb needs; the `range:` locator makes the cell revealable.
 */
export function commentsToRefs(comments: readonly ExcelComment[]): ContextRef[] {
  return comments.map((c) => ({
    id: `xl:comment:${c.id}`,
    kind: 'comment',
    surface: 'excel',
    title: `Comment by ${commentAuthor(c.authorName)}${c.address ? ` on ${c.address}` : ''}${
      c.resolved ? ' (resolved)' : ''
    }`,
    preview: c.content.slice(0, 120),
    ...(c.address ? { anchor: { matchText: c.address, locator: `range:${c.address}` } } : {}),
  }));
}

/**
 * One comment thread → a single text part: the id to reply with, its state, the cell, the comment
 * and each reply in order. Host text is JSON-quoted onto its own line (see {@link quoted}).
 */
export function commentToContext(comment: ExcelComment): ResolvedContext[] {
  const lines = [
    `Comment thread (commentId: ${comment.id}${
      comment.resolved === undefined ? '' : comment.resolved ? ', resolved' : ', open'
    })`,
    ...(comment.address ? [`On cell: ${comment.address}`] : []),
    `${quoted(commentAuthor(comment.authorName))}: ${quoted(comment.content)}`,
    ...comment.replies.map(
      (r) => `  Reply from ${quoted(commentAuthor(r.authorName))}: ${quoted(r.content)}`,
    ),
  ];
  const [ref] = commentsToRefs([comment]);
  if (!ref) return [];
  return [{ ref, value: { as: 'text', text: lines.join('\n'), mimeType: 'text/markdown' } }];
}
