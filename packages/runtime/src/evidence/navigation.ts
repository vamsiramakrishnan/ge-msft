import { evidenceHash } from './hash.js';
import type { EvidenceSnapshot } from '@ge/contracts';
import { byteLength } from '../docfs/bytes.js';
import { evidencePath } from './workspace.js';

/** Structural navigation and literal search over one immutable source revision. */
export type EvidenceQuery =
  | { kind: 'list'; parentId?: string }
  | { kind: 'grep'; literal: string; caseSensitive?: boolean; parentId?: string }
  | {
      kind: 'expand';
      id: string;
      direction: 'parent' | 'children' | 'related' | 'dependents' | 'neighbors';
    };
export interface EvidenceNode {
  id: string;
  parentId?: string;
  relatedIds?: string[];
  path: string;
  hash: string;
  textLength: number;
  /** UTF-16 offset, matching JavaScript string slicing. */
  matchOffset?: number;
}
export interface EvidenceQueryPage {
  snapshotId: string;
  nodes: EvidenceNode[];
  nextCursor?: string;
  complete: boolean;
}

export async function navigateEvidence(
  snapshot: EvidenceSnapshot,
  query: EvidenceQuery,
  options: { cursor?: string; limit?: number; maxBytes?: number } = {},
): Promise<EvidenceQueryPage> {
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
  const key = await evidenceHash(JSON.stringify([snapshot.id, query]));
  let offset = 0;
  if (options.cursor !== undefined) {
    const match = /^([a-f0-9]{64}):(\d+)$/.exec(options.cursor);
    if (!match || match[1] !== key || !Number.isSafeInteger(Number(match[2])))
      throw new Error('Cursor belongs to a different snapshot or query.');
    offset = Number(match[2]);
  }
  const units = snapshot.pages.flatMap((page) => page.units);
  const anchorIndex =
    query.kind === 'expand' ? units.findIndex((unit) => unit.id === query.id) : -1;
  if (query.kind === 'expand' && anchorIndex < 0) throw new Error('Expansion anchor is absent.');
  const anchor = units[anchorIndex];
  const matches = units.flatMap((unit, index) => {
    let matchOffset: number | undefined;
    if (query.kind === 'expand') {
      const selected =
        query.direction === 'parent'
          ? unit.id === anchor!.parentId
          : query.direction === 'children'
            ? unit.parentId === anchor!.id
            : query.direction === 'related'
              ? anchor!.relatedIds?.includes(unit.id)
              : query.direction === 'dependents'
                ? unit.relatedIds?.includes(anchor!.id)
                : Math.abs(index - anchorIndex) === 1 && unit.parentId === anchor!.parentId;
      if (!selected) return [];
    } else {
      if (query.parentId !== undefined && unit.parentId !== query.parentId) return [];
      if (query.kind === 'grep') {
        // Case folding can change string length. Locate against the original text so offsets stay valid.
        const escaped = query.literal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        matchOffset = new RegExp(escaped, query.caseSensitive ? 'u' : 'iu').exec(unit.text)?.index;
        if (matchOffset === undefined) return [];
      }
    }
    return [{ unit, index, matchOffset }];
  });
  if (offset > matches.length) throw new Error('Invalid evidence navigation offset.');
  const result: EvidenceQueryPage = { snapshotId: snapshot.id, nodes: [], complete: false };
  for (let i = offset; i < matches.length && result.nodes.length < limit; i++) {
    const { unit, index, matchOffset } = matches[i]!;
    const node: EvidenceNode = {
      id: unit.id,
      parentId: unit.parentId,
      relatedIds: unit.relatedIds,
      path: evidencePath(snapshot.id, index),
      hash: await evidenceHash(JSON.stringify(unit)),
      textLength: unit.text.length,
      matchOffset,
    };
    const candidate = { ...result, nodes: [...result.nodes, node], nextCursor: `${key}:${i + 1}` };
    if (byteLength(JSON.stringify(candidate)) > maxBytes) {
      if (!result.nodes.length) throw new Error('Node metadata exceeds the navigation budget.');
      break;
    }
    result.nodes.push(node);
  }
  const next = offset + result.nodes.length;
  result.complete = next === matches.length;
  if (!result.complete) result.nextCursor = `${key}:${next}`;
  return result;
}
