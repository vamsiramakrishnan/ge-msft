import {
  ContextProjectionRequestSchema,
  DocumentPageSchema,
  EvidenceSnapshotSchema,
  type ContextProjection,
  type ContextProjectionRequest,
  type DocumentPage,
  type EvidenceSnapshot,
  type SourceLimits,
  type SourceReceipt,
} from '@ge/contracts';
import { evidenceHash } from './hash.js';
import { navigateEvidence, type EvidenceQuery } from './navigation.js';
import { navigateEvidencePages } from './paged-navigation.js';
import { byteLength } from '../docfs/bytes.js';
import { SourceTraversal, verifySource, type EvidenceSource } from './source.js';
import { MemoryEvidenceStore, type EvidenceStore } from './store.js';

export function evidencePath(snapshot: string, index: number): string {
  return `/evidence/${snapshot}/units/${Math.floor(index / 100)}/${index}.json`;
}
export function snapshotReceipt(snapshot: EvidenceSnapshot): SourceReceipt {
  const { pages: _pages, ...receipt } = snapshot;
  return receipt;
}

/** Retrieval policy is replaceable; returned IDs must belong to the captured source. */
export type EvidenceSelector = (
  query: string,
  units: ReadonlyArray<Readonly<DocumentPage['units'][number]>>,
) => Promise<readonly string[]>;
export const lexicalEvidenceSelector: EvidenceSelector = async (query, units) => {
  const terms = [...new Set(query.toLocaleLowerCase().match(/[\p{L}\p{N}_]+/gu) ?? [])];
  return units
    .map((unit, index) => {
      const text = unit.text.toLocaleLowerCase();
      return { id: unit.id, index, score: terms.filter((term) => text.includes(term)).length };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .map((item) => item.id);
};

/** Source-neutral evidence, independent of Office, model prompts, task verbs and write authority. */
export class EvidenceWorkspace {
  constructor(
    private readonly store: EvidenceStore = new MemoryEvidenceStore(),
    private readonly selector: EvidenceSelector = lexicalEvidenceSelector,
  ) {}
  async capture(
    source: EvidenceSource,
    options: Partial<SourceLimits> & { signal?: AbortSignal } = {},
  ): Promise<SourceReceipt> {
    const { signal, ...limits } = options;
    const traversal = new SourceTraversal(source, limits);
    const pages: DocumentPage[] = [];
    for await (const page of traversal.pages(signal)) pages.push(page);
    signal?.throwIfAborted();
    const receipt = traversal.receipt;
    // Nothing is published until validated EOF. A failed/cancelled traversal leaves no partial snapshot.
    await this.store.put({ ...receipt, pages });
    return receipt;
  }
  async list(): Promise<string[]> {
    return this.store.list();
  }
  async release(id: string): Promise<void> {
    await this.store.delete(id);
  }
  async load(id: string): Promise<EvidenceSnapshot> {
    const raw = await this.store.get(id);
    if (raw === undefined) throw new Error('Evidence snapshot unavailable or released.');
    const snapshot = EvidenceSnapshotSchema.parse(raw);
    if (snapshot.id !== id) throw new Error('Evidence store returned the wrong snapshot.');
    // Treat storage as an untrusted boundary, including a caller-supplied persistent adapter.
    let index = 0;
    const traversal = new SourceTraversal(
      {
        id: snapshot.sourceId,
        readPage: async () => {
          const page = snapshot.pages[index++];
          if (!page) throw new Error('Stored evidence ended before EOF.');
          return page;
        },
      },
      { maxBytes: 256 * 1024 * 1024 },
    );
    for await (const _page of traversal.pages()) void _page;
    const { id: _id, pages: _pages, ...body } = snapshot;
    if (
      index !== snapshot.pages.length ||
      JSON.stringify(traversal.receipt.pageHashes) !== JSON.stringify(snapshot.pageHashes) ||
      traversal.receipt.unitCount !== snapshot.unitCount ||
      traversal.receipt.sourceBytes !== snapshot.sourceBytes ||
      traversal.receipt.scope !== snapshot.scope ||
      JSON.stringify(traversal.receipt.limitations) !== JSON.stringify(snapshot.limitations) ||
      (snapshot.pageUnitCounts !== undefined &&
        JSON.stringify(traversal.receipt.pageUnitCounts) !==
          JSON.stringify(snapshot.pageUnitCounts)) ||
      (await evidenceHash(JSON.stringify(body))) !== snapshot.id
    )
      throw new Error('Evidence integrity check failed.');
    return snapshot;
  }
  async verify(id: string, source: EvidenceSource, signal?: AbortSignal): Promise<void> {
    await verifySource(source, await this.receipt(id), signal);
  }
  /** Hash the small manifest, then only the page requested. A corrupt adapter cannot forge membership. */
  async receipt(id: string): Promise<SourceReceipt> {
    if (!this.store.getReceipt) return snapshotReceipt(await this.load(id));
    const raw = await this.store.getReceipt(id);
    if (!raw) throw new Error('Evidence snapshot unavailable or released.');
    const { pages: _pages, ...receipt } = EvidenceSnapshotSchema.parse({ ...raw, pages: [] });
    const { id: storedId, ...body } = receipt;
    if (storedId !== id || (await evidenceHash(JSON.stringify(body))) !== id)
      throw new Error('Evidence integrity check failed.');
    if (
      receipt.pageUnitCounts &&
      (receipt.pageUnitCounts.length !== receipt.pageHashes.length ||
        receipt.pageUnitCounts.reduce((sum, count) => sum + count, 0) !== receipt.unitCount)
    )
      throw new Error('Evidence ordinal index integrity check failed.');
    return receipt;
  }
  async page(id: string, index: number): Promise<DocumentPage> {
    if (!Number.isSafeInteger(index) || index < 0) throw new Error('Invalid evidence page.');
    if (!this.store.getPage) {
      const page = (await this.load(id)).pages[index];
      if (!page) throw new Error('Evidence page unavailable.');
      return page;
    }
    const receipt = await this.receipt(id);
    const page = DocumentPageSchema.parse(await this.store.getPage(id, index));
    if (
      index >= receipt.pageHashes.length ||
      (await evidenceHash(JSON.stringify(page))) !== receipt.pageHashes[index]
    )
      throw new Error('Evidence page integrity check failed.');
    return page;
  }
  async readAt(id: string, ordinal: number): Promise<DocumentPage['units'][number]> {
    if (!Number.isSafeInteger(ordinal) || ordinal < 0) throw new Error('Invalid evidence ordinal.');
    const receipt = await this.receipt(id);
    if (ordinal >= receipt.unitCount)
      throw new Error('Evidence unit does not exist in this snapshot.');
    if (!receipt.pageUnitCounts) {
      return structuredClone((await this.load(id)).pages.flatMap((page) => page.units)[ordinal]!);
    }
    let offset = 0;
    for (const [index, count] of receipt.pageUnitCounts.entries()) {
      if (ordinal < offset + count) {
        const page = await this.page(id, index);
        if (page.units.length !== count)
          throw new Error('Evidence ordinal index integrity check failed.');
        return structuredClone(page.units[ordinal - offset]!);
      }
      offset += count;
    }
    throw new Error('Evidence ordinal index integrity check failed.');
  }
  async read(id: string, unitId: string): Promise<DocumentPage['units'][number]> {
    if (this.store.getReceipt && this.store.getPage && this.store.locateUnit) {
      // Validate even a miss: a released or corrupt snapshot must not appear to be a valid empty one.
      await this.receipt(id);
      const index = await this.store.locateUnit(id, unitId);
      if (index === undefined) throw new Error('Evidence unit does not exist in this snapshot.');
      const unit = (await this.page(id, index)).units.find((unit) => unit.id === unitId);
      if (!unit) throw new Error('Evidence index integrity check failed.');
      return structuredClone(unit);
    }
    const unit = (await this.load(id)).pages
      .flatMap((page) => page.units)
      .find((unit) => unit.id === unitId);
    if (!unit) throw new Error('Evidence unit does not exist in this snapshot.');
    return structuredClone(unit);
  }
  async query(
    id: string,
    query: EvidenceQuery,
    options: { cursor?: string; limit?: number; maxBytes?: number } = {},
  ) {
    const receipt = await this.receipt(id);
    if (
      receipt.pageUnitCounts &&
      this.store.getPage &&
      (query.kind !== 'expand' || this.store.locateUnit)
    ) {
      let anchor: { unit: DocumentPage['units'][number]; ordinal: number } | undefined;
      if (query.kind === 'expand') {
        const index = await this.store.locateUnit!(id, query.id);
        if (index === undefined) throw new Error('Expansion anchor is absent.');
        const page = await this.page(id, index);
        const within = page.units.findIndex((unit) => unit.id === query.id);
        if (within < 0) throw new Error('Evidence index integrity check failed.');
        anchor = {
          unit: page.units[within]!,
          ordinal:
            receipt.pageUnitCounts.slice(0, index).reduce((sum, count) => sum + count, 0) + within,
        };
      }
      return navigateEvidencePages(
        receipt,
        query,
        (index) => this.page(id, index),
        anchor,
        options,
      );
    }
    return navigateEvidence(await this.load(id), query, options);
  }
  /** Exact text windows; callers can advance nextOffset without losing oversized nodes. */
  async readText(id: string, unitId: string, options: { offset?: number; maxBytes?: number } = {}) {
    const unit = await this.read(id, unitId);
    const offset = options.offset ?? 0;
    const maxBytes = options.maxBytes ?? 16 * 1024;
    if (
      !Number.isSafeInteger(offset) ||
      offset < 0 ||
      offset > unit.text.length ||
      !Number.isSafeInteger(maxBytes) ||
      maxBytes < 4 ||
      maxBytes > 96 * 1024
    )
      throw new Error('Invalid evidence text window.');
    if (
      offset > 0 &&
      /[\uDC00-\uDFFF]/.test(unit.text[offset] ?? '') &&
      /[\uD800-\uDBFF]/.test(unit.text[offset - 1]!)
    )
      throw new Error('Text offset splits a Unicode character.');
    let text = '';
    let bytes = 0;
    for (const character of unit.text.slice(offset)) {
      const size = byteLength(character);
      if (bytes + size > maxBytes) break;
      text += character;
      bytes += size;
    }
    const next = offset + text.length;
    return {
      snapshotId: id,
      id: unitId,
      offset,
      text,
      nextOffset: next < unit.text.length ? next : undefined,
      complete: next === unit.text.length,
    };
  }
  async project(id: string, input: ContextProjectionRequest = {}): Promise<ContextProjection> {
    const request = ContextProjectionRequestSchema.parse(input);
    const snapshot = await this.load(id);
    const units = snapshot.pages.flatMap((page) => page.units);
    const indexes = new Map(units.map((unit, index) => [unit.id, index]));
    const required = [...new Set(request.requiredIds)];
    for (const id of required)
      if (!indexes.has(id)) throw new Error(`Required evidence is missing: ${id}`);
    const candidates = new Set(required);
    const unresolved = new Set<string>();
    let frontier = required;
    for (let depth = 0; depth < request.relationDepth; depth++) {
      const next: string[] = [];
      for (const id of frontier) {
        const unit = units[indexes.get(id)!]!;
        for (const related of [
          ...(unit.parentId ? [unit.parentId] : []),
          ...(unit.relatedIds ?? []),
        ]) {
          if (!indexes.has(related)) {
            unresolved.add(related);
            continue;
          }
          if (!candidates.has(related)) {
            candidates.add(related);
            next.push(related);
          }
        }
      }
      frontier = next;
    }
    for (const id of required) {
      const index = indexes.get(id)!;
      for (
        let n = Math.max(0, index - request.neighbors);
        n <= Math.min(units.length - 1, index + request.neighbors);
        n++
      ) {
        if (units[n]!.parentId === units[index]!.parentId) candidates.add(units[n]!.id);
      }
    }
    const selected = await this.selector(request.query, structuredClone(units));
    for (const id of selected) {
      if (!indexes.has(id)) throw new Error('Selector returned evidence outside the snapshot.');
      candidates.add(id);
    }
    const projection: ContextProjection = {
      snapshotId: snapshot.id,
      sourceId: snapshot.sourceId,
      scope: snapshot.scope,
      sourceComplete: true,
      projectionComplete: false,
      selection:
        this.selector === lexicalEvidenceSelector
          ? 'required-structure-lexical'
          : 'required-structure-custom',
      omittedUnits: units.length,
      unresolvedRelations: unresolved.size,
      limitations: [...snapshot.limitations],
      units: [],
    };
    if (byteLength(JSON.stringify(projection)) > request.maxBytes)
      throw new Error('Projection metadata exceeds its byte budget.');
    const requiredSet = new Set(required);
    let encodedBytes = byteLength(JSON.stringify(projection));
    for (const id of candidates) {
      const index = indexes.get(id)!;
      const unit = units[index]!;
      const entry = {
        id,
        hash: '0'.repeat(64),
        path: evidencePath(snapshot.id, index),
        text: unit.text,
      };
      const nextOmitted = projection.omittedUnits - 1;
      const delta =
        byteLength(JSON.stringify(entry)) +
        (projection.units.length ? 1 : 0) +
        String(nextOmitted).length -
        String(projection.omittedUnits).length +
        (nextOmitted === 0 ? -1 : 0); // JSON true is one byte shorter than false.
      if (encodedBytes + delta > request.maxBytes) {
        if (requiredSet.has(id))
          throw new Error('Required evidence exceeds the projection budget; it was not truncated.');
        continue;
      }
      entry.hash = await evidenceHash(JSON.stringify(unit));
      projection.units.push(entry);
      projection.omittedUnits = nextOmitted;
      projection.projectionComplete = nextOmitted === 0;
      encodedBytes += delta;
    }
    return projection;
  }
}
