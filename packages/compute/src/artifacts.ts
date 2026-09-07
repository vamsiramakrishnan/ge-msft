import {
  TableArtifactSchema,
  contentHash,
  sourceVersion,
  type ArtifactSummary,
  type CellSnapshot,
  type CellValue,
  type TableArtifact,
} from '@ge/contracts';

export const MAX_ARTIFACT_BYTES = 16 * 1024 * 1024;
export const MAX_WORKSPACE_BYTES = 32 * 1024 * 1024;
const bytes = (value: unknown): number =>
  new TextEncoder().encode(JSON.stringify(value)).byteLength;
export function columnType(values: CellValue[]): TableArtifact['columns'][number]['type'] {
  const types = new Set(values.filter((v) => v !== null && v !== '').map((v) => typeof v));
  return types.size === 0
    ? 'empty'
    : types.size > 1
      ? 'mixed'
      : ([...types][0] as 'string' | 'number' | 'boolean');
}
export interface ArtifactInput {
  title: string;
  labels: string[];
  rows: CellValue[][];
  sources: TableArtifact['sources'];
  lineage: TableArtifact['lineage'];
  truncated?: boolean;
}

/** Immutable, content-addressed artifacts; no implicit overwrite or unbounded history. */
export class ArtifactStore {
  private readonly entries = new Map<string, TableArtifact>();
  private totalBytes = 0;
  async add(input: ArtifactInput): Promise<TableArtifact> {
    if (bytes(input) > MAX_ARTIFACT_BYTES)
      throw new Error('Artifact exceeds the 16 MiB budget. Narrow the source or result.');
    const columns = input.labels.map((label, i) => ({
      name: `c${i}`,
      label: label.slice(0, 512),
      type: columnType(input.rows.map((r) => r[i] ?? null)),
    }));
    const body = {
      columns,
      rows: input.rows,
      sources: input.sources,
      lineage: input.lineage,
      truncated: input.truncated ?? false,
    };
    const hash = await contentHash(body);
    const id = `a_${hash.slice(7, 31)}`;
    const existing = this.entries.get(id);
    if (existing) return structuredClone(existing);
    const artifact = TableArtifactSchema.parse({
      ...body,
      id,
      hash,
      title: input.title,
      createdAt: new Date().toISOString(),
    });
    const size = bytes(artifact);
    if (this.entries.size >= 32 || this.totalBytes + size > MAX_WORKSPACE_BYTES)
      throw new Error('Analysis workspace is full. Remove an artifact before adding another.');
    this.entries.set(id, structuredClone(artifact));
    this.totalBytes += size;
    return structuredClone(artifact);
  }
  async fromSnapshot(snapshot: CellSnapshot, hasHeaders = true): Promise<TableArtifact> {
    const labels = snapshot.values[0]!.map((v, i) =>
      hasHeaders ? String(v ?? '').trim() || `Column ${i + 1}` : `Column ${i + 1}`,
    );
    return this.add({
      title: snapshot.locator,
      labels,
      rows: snapshot.values.slice(hasHeaders ? 1 : 0),
      sources: [sourceVersion(snapshot)],
      lineage: { parents: [], operation: 'snapshot' },
    });
  }
  get(id: string): TableArtifact {
    const artifact = this.entries.get(id);
    if (!artifact) throw new Error('This artifact is no longer available. Capture it again.');
    return structuredClone(artifact);
  }
  list(): ArtifactSummary[] {
    return [...this.entries.values()].map(({ rows, ...meta }) => ({
      ...structuredClone(meta),
      rowCount: rows.length,
      preview: structuredClone(rows.slice(0, 8)),
    }));
  }
  remove(id: string): void {
    const entry = this.entries.get(id);
    if (entry) {
      this.totalBytes -= bytes(entry);
      this.entries.delete(id);
    }
  }
  clear(): void {
    this.entries.clear();
    this.totalBytes = 0;
  }
}
