import type { DocumentPage, SourceReceipt } from '@ge/contracts';
import { byteLength } from '../docfs/bytes.js';
import { evidenceHash } from './hash.js';
import { evidencePath } from './workspace.js';
import type { EvidenceNode, EvidenceQuery, EvidenceQueryPage } from './navigation.js';

/** At most eight stored pages per call. Empty result pages still carry a progressing cursor. */
export async function navigateEvidencePages(
  receipt: SourceReceipt,
  query: EvidenceQuery,
  readPage: (index: number) => Promise<DocumentPage>,
  anchor: { unit: DocumentPage['units'][number]; ordinal: number } | undefined,
  options: { cursor?: string; limit?: number; maxBytes?: number } = {},
): Promise<EvidenceQueryPage> {
  const counts = receipt.pageUnitCounts;
  if (!counts || counts.length !== receipt.pageHashes.length)
    throw new Error('Authenticated page counts are required.');
  const limit = options.limit ?? 50;
  const maxBytes = options.maxBytes ?? 16 * 1024;
  if (
    !Number.isSafeInteger(limit) ||
    limit < 1 ||
    limit > 100 ||
    !Number.isSafeInteger(maxBytes) ||
    maxBytes < 512 ||
    maxBytes > 96 * 1024
  )
    throw new Error('Invalid evidence navigation budget.');
  if (query.kind === 'grep' && (!query.literal || query.literal.length > 8192))
    throw new Error('Search requires a bounded nonempty literal.');
  if (query.kind === 'expand' && !anchor) throw new Error('Expansion anchor is absent.');
  const key = await evidenceHash(JSON.stringify([receipt.id, query, 'paged-v1']));
  let pageIndex = 0;
  let unitIndex = 0;
  if (options.cursor !== undefined) {
    const match = /^([a-f0-9]{64}):(\d+):(\d+)$/.exec(options.cursor);
    if (
      !match ||
      match[1] !== key ||
      !Number.isSafeInteger(Number(match[2])) ||
      !Number.isSafeInteger(Number(match[3]))
    )
      throw new Error('Cursor belongs to a different snapshot or query.');
    pageIndex = Number(match[2]);
    unitIndex = Number(match[3]);
    if (pageIndex > counts.length || unitIndex > (counts[pageIndex] ?? 0))
      throw new Error('Invalid evidence navigation offset.');
  }
  const result: EvidenceQueryPage = { snapshotId: receipt.id, nodes: [], complete: false };
  let ordinal = counts.slice(0, pageIndex).reduce((sum, count) => sum + count, 0);
  let scanned = 0;
  const literal =
    query.kind === 'grep'
      ? new RegExp(
          query.literal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
          query.caseSensitive ? 'u' : 'iu',
        )
      : undefined;
  while (pageIndex < counts.length && scanned++ < 8) {
    const page = await readPage(pageIndex);
    if (page.units.length !== counts[pageIndex])
      throw new Error('Evidence ordinal index integrity check failed.');
    for (; unitIndex < page.units.length; unitIndex++) {
      const unit = page.units[unitIndex]!;
      const index = ordinal + unitIndex;
      let matchOffset: number | undefined;
      let selected = true;
      if (query.kind === 'expand') {
        selected =
          query.direction === 'parent'
            ? unit.id === anchor!.unit.parentId
            : query.direction === 'children'
              ? unit.parentId === anchor!.unit.id
              : query.direction === 'related'
                ? Boolean(anchor!.unit.relatedIds?.includes(unit.id))
                : query.direction === 'dependents'
                  ? Boolean(unit.relatedIds?.includes(anchor!.unit.id))
                  : Math.abs(index - anchor!.ordinal) === 1 &&
                    unit.parentId === anchor!.unit.parentId;
      } else {
        selected = query.parentId === undefined || query.parentId === unit.parentId;
        if (selected && literal) {
          matchOffset = literal.exec(unit.text)?.index;
          selected = matchOffset !== undefined;
        }
      }
      if (!selected) continue;
      const node: EvidenceNode = {
        id: unit.id,
        parentId: unit.parentId,
        relatedIds: unit.relatedIds,
        path: evidencePath(receipt.id, index),
        hash: await evidenceHash(JSON.stringify(unit)),
        textLength: unit.text.length,
        ...(matchOffset !== undefined ? { matchOffset } : {}),
      };
      const candidate = {
        ...result,
        nodes: [...result.nodes, node],
        nextCursor: `${key}:${pageIndex}:${unitIndex + 1}`,
      };
      if (result.nodes.length >= limit || byteLength(JSON.stringify(candidate)) > maxBytes) {
        if (!result.nodes.length) throw new Error('Node metadata exceeds the navigation budget.');
        result.nextCursor = `${key}:${pageIndex}:${unitIndex}`;
        return result;
      }
      result.nodes.push(node);
    }
    ordinal += page.units.length;
    pageIndex++;
    unitIndex = 0;
  }
  result.complete = pageIndex === counts.length;
  if (!result.complete) result.nextCursor = `${key}:${pageIndex}:${unitIndex}`;
  return result;
}
