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
  const headerIdx = values.findIndex((row) => row.some((cell) => cell.trim() !== ''));
  if (headerIdx === -1) return { columns: [], rows: [] };
  return { columns: values[headerIdx]!, rows: values.slice(headerIdx + 1) };
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
