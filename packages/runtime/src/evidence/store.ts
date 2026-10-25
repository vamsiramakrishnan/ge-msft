import type { DocumentPage, EvidenceSnapshot, SourceReceipt } from '@ge/contracts';
import { byteLength } from '../docfs/bytes.js';

/** Atomic publication port. Implementations must isolate principals and enforce aggregate quotas.
 * Persist only source data; never credentials, approval grants, or executable callbacks. */
export interface EvidenceStore {
  get(id: string): Promise<unknown | undefined>;
  put(snapshot: EvidenceSnapshot): Promise<void>;
  delete(id: string): Promise<void>;
  list(): Promise<string[]>;
  /** Optional indexed port. Every returned receipt/page is independently validated by the reader. */
  getReceipt?(id: string): Promise<SourceReceipt | undefined>;
  getPage?(id: string, page: number): Promise<DocumentPage | undefined>;
  locateUnit?(id: string, unitId: string): Promise<number | undefined>;
}

/** No implicit eviction: an outstanding reference must not silently point at different evidence. */
export class MemoryEvidenceStore implements EvidenceStore {
  private readonly entries = new Map<
    string,
    { receipt: string; pages: string[]; units: Map<string, number>; bytes: number }
  >();
  private bytes = 0;
  constructor(
    private readonly maxBytes = 128 * 1024 * 1024,
    private readonly maxSnapshots = 8,
  ) {
    if (
      !Number.isSafeInteger(maxBytes) ||
      maxBytes < 1 ||
      !Number.isSafeInteger(maxSnapshots) ||
      maxSnapshots < 1
    )
      throw new Error('Evidence store quotas must be positive safe integers.');
  }
  async get(id: string): Promise<unknown | undefined> {
    const entry = this.entries.get(id);
    return entry === undefined
      ? undefined
      : { ...JSON.parse(entry.receipt), pages: entry.pages.map((page) => JSON.parse(page)) };
  }
  async put(snapshot: EvidenceSnapshot): Promise<void> {
    const { pages, ...receipt } = snapshot;
    const encoded = JSON.stringify({ ...receipt, pages });
    const prior = await this.get(snapshot.id);
    if (prior !== undefined && JSON.stringify(prior) !== encoded)
      throw new Error('Immutable snapshot conflict.');
    if (prior !== undefined) return;
    const bytes = byteLength(encoded);
    if (this.entries.size >= this.maxSnapshots || this.bytes + bytes > this.maxBytes)
      throw new Error('Evidence store quota exceeded. Release an unused snapshot explicitly.');
    this.entries.set(snapshot.id, {
      bytes,
      receipt: JSON.stringify(receipt),
      pages: pages.map((page) => JSON.stringify(page)),
      units: new Map(
        pages.flatMap((page, index) => page.units.map((unit) => [unit.id, index] as const)),
      ),
    });
    this.bytes += bytes;
  }
  async delete(id: string): Promise<void> {
    const prior = this.entries.get(id);
    if (prior !== undefined) {
      this.bytes -= prior.bytes;
      this.entries.delete(id);
    }
  }
  async list(): Promise<string[]> {
    return [...this.entries.keys()];
  }
  async getReceipt(id: string): Promise<SourceReceipt | undefined> {
    const receipt = this.entries.get(id)?.receipt;
    return receipt === undefined ? undefined : JSON.parse(receipt);
  }
  async getPage(id: string, page: number): Promise<DocumentPage | undefined> {
    const encoded = this.entries.get(id)?.pages[page];
    return encoded === undefined ? undefined : JSON.parse(encoded);
  }
  async locateUnit(id: string, unitId: string): Promise<number | undefined> {
    return this.entries.get(id)?.units.get(unitId);
  }
}
