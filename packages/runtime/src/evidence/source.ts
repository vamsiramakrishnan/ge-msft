import {
  DocumentPageSchema,
  SourceLimitsSchema,
  type DocumentPage,
  type DocumentReadOptions,
  type SourceLimits,
  type SourceReceipt,
} from '@ge/contracts';
import { evidenceHash } from './hash.js';
import { byteLength } from '../docfs/bytes.js';

/** Caller-owned identity; one instance must never span users or authorization scopes. */
export interface EvidenceSource {
  id: string;
  readPage(options: DocumentReadOptions): Promise<DocumentPage>;
}

/** One traversal owns coverage. No receipt exists until EOF; early return is never completion. */
export class SourceTraversal {
  readonly limits: SourceLimits;
  private result?: SourceReceipt;
  private started = false;
  constructor(
    readonly source: EvidenceSource,
    limits: Partial<SourceLimits> = {},
  ) {
    if (!source.id || source.id.length > 1000) throw new Error('Invalid evidence source identity.');
    this.limits = SourceLimitsSchema.parse(limits);
  }
  get receipt(): SourceReceipt {
    if (!this.result) throw new Error('Source traversal has not reached EOF.');
    return structuredClone(this.result);
  }
  async *pages(signal?: AbortSignal): AsyncGenerator<DocumentPage> {
    if (this.started) throw new Error('A source traversal is single-use.');
    this.started = true;
    let cursor: string | undefined;
    const cursors = new Set<string>();
    const ids = new Set<string>();
    const pageHashes: string[] = [];
    const pageUnitCounts: number[] = [];
    const limitations = new Set<string>();
    let scope = '';
    let sourceBytes = 0;
    do {
      signal?.throwIfAborted();
      if (pageHashes.length >= this.limits.maxPages)
        throw new Error('Source page limit reached before EOF.');
      const key = cursor ?? '<start>';
      if (cursors.has(key))
        throw new Error('Document reader repeated a cursor; coverage is incomplete.');
      cursors.add(key);
      const page = DocumentPageSchema.parse(await this.source.readPage({ cursor, signal }));
      signal?.throwIfAborted();
      if (scope && scope !== page.scope) throw new Error('Document scope changed during the read.');
      scope = page.scope;
      for (const limitation of page.limitations) limitations.add(limitation);
      if (limitations.size > 40) throw new Error('Too many extraction limitations.');
      const encoded = JSON.stringify(page);
      sourceBytes += byteLength(encoded);
      if (sourceBytes > this.limits.maxBytes)
        throw new Error('Source-byte budget exceeded; coverage is incomplete.');
      for (const unit of page.units) {
        if (ids.has(unit.id))
          throw new Error('Document reader repeated a source unit; coverage is ambiguous.');
        ids.add(unit.id);
        if (ids.size > this.limits.maxUnits) throw new Error('Source-unit limit reached.');
      }
      pageHashes.push(await evidenceHash(encoded));
      pageUnitCounts.push(page.units.length);
      cursor = page.nextCursor;
      yield page;
    } while (cursor !== undefined);
    signal?.throwIfAborted();
    const body = {
      sourceId: this.source.id,
      scope,
      pageHashes,
      unitCount: ids.size,
      sourceBytes,
      limitations: [...limitations],
      complete: true as const,
      pageUnitCounts,
    };
    this.result = { ...body, id: await evidenceHash(JSON.stringify(body)) };
  }
}

/** Full-source validation also catches newly added evidence, not merely changed selected units. */
export async function verifySource(
  source: EvidenceSource,
  receipt: SourceReceipt,
  signal?: AbortSignal,
): Promise<void> {
  if (source.id !== receipt.sourceId) throw new Error('Evidence belongs to a different source.');
  const traversal = new SourceTraversal(source, { maxBytes: 256 * 1024 * 1024 });
  let page = 0;
  for await (const value of traversal.pages(signal)) {
    if ((await evidenceHash(JSON.stringify(value))) !== receipt.pageHashes[page++])
      throw new Error(
        'The document changed since capture. Refresh the evidence before continuing.',
      );
  }
  const { id: _id, ...body } = traversal.receipt;
  if (receipt.pageUnitCounts === undefined) delete body.pageUnitCounts;
  if ((await evidenceHash(JSON.stringify(body))) !== receipt.id)
    throw new Error('Document coverage changed since capture.');
}
