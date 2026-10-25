import type { DocumentPage, SourceReceipt } from '@ge/contracts';
import { byteLength } from '../docfs/bytes.js';
import { evidenceHash } from './hash.js';

interface IndexedNode {
  id: string;
  hash: string;
  links: readonly string[];
  complete: boolean;
}
interface IndexedPage {
  nodes: readonly IndexedNode[];
  bytes: number;
  edges: number;
}
interface RevisionIndex {
  receipt: SourceReceipt;
  pages: Map<number, IndexedPage>;
  nodes: Map<string, IndexedNode>;
  reverse: Map<string, Set<string>>;
  nextPage: number;
  bytes: number;
  edges: number;
  limited: boolean;
  reusedPages: number;
}
export interface DependencyIndexProgress {
  snapshotId: string;
  indexedPages: number;
  totalPages: number;
  indexedUnits: number;
  omittedUnits: number;
  unresolvedLinks: number;
  incompleteRelationUnits: number;
  reusedPages: number;
  complete: boolean;
  limited: boolean;
  /** Always revision-local captured links, never host execution or freshness authority. */
  coverage: 'captured-links';
  limitations: string[];
}
export interface DependencyImpact {
  previousSnapshotId: string;
  snapshotId: string;
  changedIds: string[];
  affectedIds: string[];
  complete: boolean;
  requiresFullRefresh: boolean;
  limitations: string[];
}

const COVERAGE_LIMIT =
  'Only captured parent/formula links are indexed. Dependency impact is navigation data, not a freshness check, approval or write target.';

/**
 * An optional, bounded accelerator over authenticated evidence pages. Publication of a source
 * snapshot does not depend on indexing it. Each call reads at most eight pages; identical page
 * hashes are reused across revisions. A partial index never certifies selective invalidation.
 */
export class EvidenceDependencyIndex {
  private readonly revisions = new Map<string, RevisionIndex>();
  private readonly sharedPages = new Map<string, { page: IndexedPage; references: number }>();
  private readonly pending = new Map<string, Promise<unknown>>();
  constructor(
    private readonly limits = {
      maxRevisions: 4,
      maxUnits: 16_000,
      maxEdges: 64_000,
      maxBytes: 4 * 1024 * 1024,
    },
  ) {
    if (Object.values(limits).some((value) => !Number.isSafeInteger(value) || value < 1))
      throw new Error('Invalid dependency index limits.');
  }
  async release(id: string): Promise<void> {
    await this.pending.get(id)?.catch(() => {});
    const revision = this.revisions.get(id);
    if (!revision) return;
    for (const index of revision.pages.keys()) {
      const hash = revision.receipt.pageHashes[index]!;
      const shared = this.sharedPages.get(hash)!;
      if (--shared.references === 0) this.sharedPages.delete(hash);
    }
    this.revisions.delete(id);
  }
  async advance(
    receipt: SourceReceipt,
    readPage: (index: number) => Promise<DocumentPage>,
    options: { maxPages?: number; signal?: AbortSignal } = {},
  ): Promise<DependencyIndexProgress> {
    const previous = this.pending.get(receipt.id) ?? Promise.resolve();
    const operation = previous
      .catch(() => {})
      .then(() => this.advanceRevision(receipt, readPage, options));
    this.pending.set(receipt.id, operation);
    try {
      return await operation;
    } finally {
      if (this.pending.get(receipt.id) === operation) this.pending.delete(receipt.id);
    }
  }
  private async advanceRevision(
    receipt: SourceReceipt,
    readPage: (index: number) => Promise<DocumentPage>,
    options: { maxPages?: number; signal?: AbortSignal },
  ): Promise<DependencyIndexProgress> {
    const maxPages = options.maxPages ?? 8;
    if (!Number.isSafeInteger(maxPages) || maxPages < 1 || maxPages > 8)
      throw new Error('Index at most eight evidence pages per call.');
    options.signal?.throwIfAborted();
    let revision = this.revisions.get(receipt.id);
    if (!revision) {
      if (this.revisions.size >= this.limits.maxRevisions)
        throw new Error('Dependency index quota exceeded. Release an unused revision explicitly.');
      revision = {
        receipt: structuredClone(receipt),
        pages: new Map(),
        nodes: new Map(),
        reverse: new Map(),
        nextPage: 0,
        bytes: 0,
        edges: 0,
        limited: false,
        reusedPages: 0,
      };
      this.revisions.set(receipt.id, revision);
    }
    if (JSON.stringify(revision.receipt) !== JSON.stringify(receipt))
      throw new Error('Dependency index revision conflict.');
    let scanned = 0;
    while (
      !revision.limited &&
      revision.nextPage < receipt.pageHashes.length &&
      scanned++ < maxPages
    ) {
      options.signal?.throwIfAborted();
      const pageIndex = revision.nextPage;
      const hash = receipt.pageHashes[pageIndex]!;
      let shared = this.sharedPages.get(hash);
      const reused = Boolean(shared);
      if (!shared) {
        const page = await readPage(pageIndex);
        options.signal?.throwIfAborted();
        // Defence in depth: an SDK caller cannot seed the reusable index with a forged page.
        if ((await evidenceHash(JSON.stringify(page))) !== hash)
          throw new Error('Dependency page integrity check failed.');
        const nodes: IndexedNode[] = [];
        for (const unit of page.units)
          nodes.push({
            id: unit.id,
            hash: await evidenceHash(JSON.stringify(unit)),
            links: [
              ...new Set([...(unit.parentId ? [unit.parentId] : []), ...(unit.relatedIds ?? [])]),
            ],
            complete: unit.relationsComplete === true,
          });
        shared = {
          page: {
            nodes,
            bytes: byteLength(JSON.stringify(nodes)),
            edges: nodes.reduce((sum, node) => sum + node.links.length, 0),
          },
          references: 0,
        };
      }
      const page = shared.page;
      if (
        revision.nodes.size + page.nodes.length > this.limits.maxUnits ||
        revision.edges + page.edges > this.limits.maxEdges ||
        revision.bytes + page.bytes > this.limits.maxBytes
      ) {
        revision.limited = true;
        break;
      }
      if (page.nodes.some((node) => revision!.nodes.has(node.id)))
        throw new Error('Duplicate dependency unit in source revision.');
      if (!this.sharedPages.has(hash)) this.sharedPages.set(hash, shared);
      shared.references++;
      revision.pages.set(pageIndex, page);
      revision.bytes += page.bytes;
      revision.edges += page.edges;
      if (reused) revision.reusedPages++;
      for (const node of page.nodes) {
        revision.nodes.set(node.id, node);
        for (const link of node.links) {
          const dependents = revision.reverse.get(link) ?? new Set<string>();
          dependents.add(node.id);
          revision.reverse.set(link, dependents);
        }
      }
      revision.nextPage++;
    }
    return this.progress(revision);
  }
  /** Both revisions must have been advanced explicitly; this method performs no hidden source reads. */
  impact(
    previousId: string,
    id: string,
    options: { maxNodes?: number; maxDepth?: number } = {},
  ): DependencyImpact {
    const previous = this.revisions.get(previousId);
    const current = this.revisions.get(id);
    if (!previous || !current)
      throw new Error('Index both revisions before comparing dependencies.');
    if (
      previous.receipt.sourceId !== current.receipt.sourceId ||
      previous.receipt.scope !== current.receipt.scope
    )
      throw new Error('Dependency revisions belong to different sources or scopes.');
    const maxNodes = options.maxNodes ?? 256;
    const maxDepth = options.maxDepth ?? 8;
    if (
      !Number.isSafeInteger(maxNodes) ||
      maxNodes < 1 ||
      maxNodes > 1000 ||
      !Number.isSafeInteger(maxDepth) ||
      maxDepth < 0 ||
      maxDepth > 32
    )
      throw new Error('Invalid dependency impact budget.');
    const p = this.progress(previous);
    const c = this.progress(current);
    const result: DependencyImpact = {
      previousSnapshotId: previousId,
      snapshotId: id,
      changedIds: [],
      affectedIds: [],
      complete:
        p.complete &&
        c.complete &&
        p.unresolvedLinks === 0 &&
        c.unresolvedLinks === 0 &&
        p.incompleteRelationUnits === 0 &&
        c.incompleteRelationUnits === 0,
      requiresFullRefresh: false,
      limitations: [...new Set([...p.limitations, ...c.limitations])],
    };
    const changed = new Set<string>();
    for (const [unitId, node] of previous.nodes)
      if (current.nodes.get(unitId)?.hash !== node.hash) changed.add(unitId);
    for (const unitId of current.nodes.keys()) if (!previous.nodes.has(unitId)) changed.add(unitId);
    result.changedIds = [...changed].slice(0, maxNodes);
    if (changed.size > maxNodes) result.complete = false;
    const affected = new Set(result.changedIds);
    let frontier = result.changedIds;
    let depth = 0;
    while (frontier.length) {
      const next: string[] = [];
      for (const unitId of frontier) {
        const dependents = new Set([
          ...(previous.reverse.get(unitId) ?? []),
          ...(current.reverse.get(unitId) ?? []),
        ]);
        for (const dependent of dependents) {
          if (affected.has(dependent)) continue;
          if (depth >= maxDepth || affected.size >= maxNodes) {
            result.complete = false;
            continue;
          }
          affected.add(dependent);
          next.push(dependent);
        }
      }
      frontier = next;
      depth++;
    }
    result.affectedIds = [...affected];
    result.requiresFullRefresh = !result.complete;
    if (!result.complete)
      result.limitations.push(
        'The dependency impact is incomplete. Refresh the complete source; do not treat omitted units as unaffected.',
      );
    return result;
  }
  private progress(revision: RevisionIndex): DependencyIndexProgress {
    let unresolvedLinks = 0;
    for (const id of revision.reverse.keys()) if (!revision.nodes.has(id)) unresolvedLinks++;
    let incompleteRelationUnits = 0;
    for (const node of revision.nodes.values()) if (!node.complete) incompleteRelationUnits++;
    return {
      snapshotId: revision.receipt.id,
      indexedPages: revision.nextPage,
      totalPages: revision.receipt.pageHashes.length,
      indexedUnits: revision.nodes.size,
      omittedUnits: Math.max(0, revision.receipt.unitCount - revision.nodes.size),
      unresolvedLinks,
      incompleteRelationUnits,
      reusedPages: revision.reusedPages,
      complete: !revision.limited && revision.nextPage === revision.receipt.pageHashes.length,
      limited: revision.limited,
      coverage: 'captured-links',
      limitations: [...revision.receipt.limitations, COVERAGE_LIMIT],
    };
  }
}
