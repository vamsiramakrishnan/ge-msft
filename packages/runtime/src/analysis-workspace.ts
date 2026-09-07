import { z } from 'zod';
import {
  ArtifactStore,
  reconciliationQuery,
  ReconciliationSpecSchema,
  assertExactDecimalColumn,
  type ComputeEngine,
} from '@ge/compute';
import {
  asChangeId,
  type ActuationRequest,
  type ArtifactSummary,
  type TableArtifact,
} from '@ge/contracts';
import type { DocBridge } from './bridge.js';

export const AnalysisActionSchema = z.discriminatedUnion('kind', [
  z.object({
    kind: z.literal('capture'),
    range: z.string().min(1).max(1024),
    headers: z.boolean().default(true),
  }),
  z.object({
    kind: z.literal('query'),
    sql: z.string().min(1).max(32768),
    inputs: z.array(z.string()).min(1).max(16),
    title: z.string().min(1).max(256).default('Query result'),
    requiredColumns: z
      .array(
        z
          .object({
            input: z.string().min(1),
            indices: z.array(z.number().int().nonnegative().max(16383)).min(1).max(64),
            exactDecimal: z.boolean().optional(),
          })
          .strict(),
      )
      .max(16)
      .optional(),
  }),
  z.object({ kind: z.literal('reconcile'), spec: ReconciliationSpecSchema }),
  z.object({ kind: z.literal('inspect'), id: z.string() }),
  z.object({ kind: z.literal('remove'), id: z.string() }),
  z.object({
    kind: z.literal('filter'),
    id: z.string(),
    status: z.enum(['matched', 'variance', 'unpaid', 'unallocated', 'invalid']),
  }),
  z.object({
    kind: z.literal('materialize'),
    id: z.string(),
    destination: z.string().min(1).max(1024),
    whenNonEmpty: z.boolean().optional(),
  }),
  z.object({ kind: z.literal('recovery') }),
  z.object({ kind: z.literal('undo'), id: z.string() }),
  z.object({ kind: z.literal('resume'), id: z.string() }),
  z.object({ kind: z.literal('forget'), id: z.string() }),
]);
export type AnalysisAction = z.input<typeof AnalysisActionSchema>;
export interface ActionOffer {
  id: string;
  title: string;
  detail: string;
  count: number;
  action: AnalysisAction;
  artifactId: string;
  approval: 'none';
}
export interface AnalysisState {
  artifacts: ArtifactSummary[];
  selected?: string;
  offers: ActionOffer[];
  note?: string;
}

/** All routes share this service: typed UI actions and the model's constrained analysis command. */
export class AnalysisWorkspace {
  readonly artifacts = new ArtifactStore();
  private engine?: ComputeEngine;
  private selected?: string;
  constructor(
    private readonly bridge: DocBridge,
    private readonly createEngine: () => Promise<ComputeEngine>,
  ) {}
  async fresh(tables: readonly TableArtifact[]): Promise<void> {
    if (!this.bridge.captureCells)
      throw new Error('This host does not support versioned cell analysis.');
    const sources = new Map(
      tables.flatMap((t) => t.sources).map((s) => [`${s.documentId}:${s.locator}`, s]),
    );
    // Different versions of one source cannot participate in a single trusted result.
    for (const table of tables)
      for (const source of table.sources)
        if (sources.get(`${source.documentId}:${source.locator}`)?.hash !== source.hash)
          throw new Error(
            'The inputs contain different versions of the same source. Capture them again.',
          );
    for (const source of sources.values()) {
      const current = await this.bridge.captureCells(source.locator);
      if (current.hash !== source.hash || current.documentId !== source.documentId)
        throw new Error(`Source changed: ${source.locator}. Capture and analyze it again.`);
    }
  }
  async execute(raw: AnalysisAction, signal?: AbortSignal): Promise<TableArtifact | undefined> {
    const action = AnalysisActionSchema.parse(raw);
    signal?.throwIfAborted();
    if (action.kind === 'remove') {
      this.artifacts.remove(action.id);
      if (this.selected === action.id) this.selected = undefined;
      return;
    }
    if (action.kind === 'inspect') {
      const artifact = this.artifacts.get(action.id);
      await this.fresh([artifact]);
      this.selected = artifact.id;
      return artifact;
    }
    let output: TableArtifact;
    if (action.kind === 'capture') {
      if (!this.bridge.captureCells)
        throw new Error('Versioned cell analysis is unavailable in this host.');
      output = await this.artifacts.fromSnapshot(
        await this.bridge.captureCells(action.range),
        action.headers,
      );
    } else {
      let sql: string;
      let tables: TableArtifact[];
      let title: string;
      let operation: 'query' | 'reconcile' = 'query';
      if (action.kind === 'query') {
        sql = action.sql;
        tables = action.inputs.map((id) => this.artifacts.get(id));
        title = action.title;
        for (const required of action.requiredColumns ?? []) {
          const table = tables.find((input) => input.id === required.input);
          if (!table) throw new Error('A required column belongs to an undeclared query input.');
          for (const index of required.indices) {
            if (!table.columns[index])
              throw new Error(
                `Column ${index + 1} (index ${index}) does not exist in ${table.title}; the captured range has ${table.columns.length} columns. Expand the range or choose another column.`,
              );
            if (required.exactDecimal) assertExactDecimalColumn(table, index);
          }
        }
      } else if (action.kind === 'reconcile') {
        tables = [this.artifacts.get(action.spec.left), this.artifacts.get(action.spec.right)];
        sql = reconciliationQuery(tables[0]!, tables[1]!, action.spec);
        title = 'Invoice reconciliation';
        operation = 'reconcile';
      } else if (action.kind === 'filter') {
        const table = this.artifacts.get(action.id);
        const status = table.columns.find((c) => c.label === 'status');
        if (!status || table.lineage.operation !== 'reconcile')
          throw new Error('This finding no longer belongs to a reconciliation result.');
        tables = [table];
        sql = `SELECT * FROM ${table.id} WHERE ${status.name} = '${action.status}'`;
        title = `${action.status} findings`;
      } else throw new Error('This operation requires the session approval or recovery service.');
      await this.fresh(tables);
      signal?.throwIfAborted();
      this.engine ??= await this.createEngine();
      const result = await this.engine.query(sql, tables, signal);
      await this.fresh(tables);
      signal?.throwIfAborted();
      const sources = [
        ...new Map(
          tables.flatMap((t) => t.sources).map((s) => [`${s.documentId}:${s.locator}`, s]),
        ).values(),
      ];
      output = await this.artifacts.add({
        title,
        labels: action.kind === 'filter' ? tables[0]!.columns.map((c) => c.label) : result.columns,
        rows: result.rows,
        sources,
        lineage: { parents: tables.map((t) => t.id), operation, expression: sql },
        truncated: result.truncated || tables.some((t) => t.truncated),
      });
    }
    signal?.throwIfAborted();
    this.selected = output.id;
    return output;
  }
  async materialize(id: string, destination: string): Promise<ActuationRequest> {
    const artifact = this.artifacts.get(id);
    await this.fresh([artifact]);
    if (artifact.truncated)
      throw new Error('This result is truncated. Narrow the query before writing it.');
    if (!artifact.rows.length) throw new Error('This result has no rows to write.');
    return {
      changeId: asChangeId(`analysis:${crypto.randomUUID()}`),
      surface: this.bridge.surface,
      kind: 'write-cells',
      params: {
        target: { range: destination },
        cellValues: [artifact.columns.map((c) => c.label), ...artifact.rows],
      },
      preconditions: artifact.sources,
    };
  }
  /** Model tool receipts contain handles and bounded excerpts, never the whole UI workspace. */
  receipt(): unknown {
    const state = this.state();
    const selected = state.artifacts.find((a) => a.id === state.selected);
    const result = selected
      ? {
          id: selected.id,
          rowCount: selected.rowCount,
          truncated: selected.truncated,
          columns: selected.columns.map((c) => ({ ...c, label: c.label.slice(0, 64) })),
          sources: selected.sources.map((s) => ({
            locator: s.locator.slice(0, 128),
            hash: s.hash,
          })),
          preview: {
            columns: selected.columns.slice(0, 16).map((c) => c.name),
            rows: selected.preview.map((row) =>
              row.slice(0, 16).map((v) => (typeof v === 'string' ? v.slice(0, 128) : v)),
            ),
          },
          previewLimited: true,
        }
      : undefined;
    const receipt = {
      artifacts: state.artifacts.map((a) => ({
        id: a.id,
        title: a.title.slice(0, 128),
        rows: a.rowCount,
        columns: a.columns.length,
        truncated: a.truncated,
      })),
      result,
      findings: state.offers.map((o) => ({ title: o.title, count: o.count, action: o.action })),
    };
    if (result && new TextEncoder().encode(JSON.stringify(receipt)).byteLength > 64 * 1024) {
      result.preview.rows = [];
      result.columns = result.columns.map((c) => ({ ...c, label: c.label.slice(0, 16) }));
    }
    return receipt;
  }

  state(): AnalysisState {
    const artifacts = this.artifacts.list();
    const selected = artifacts.find((a) => a.id === this.selected);
    const offers: ActionOffer[] = [];
    if (selected?.lineage.operation === 'reconcile') {
      const table = this.artifacts.get(selected.id);
      const index = table.columns.findIndex((c) => c.label === 'status');
      for (const status of ['invalid', 'variance', 'unpaid', 'unallocated', 'matched'] as const) {
        const count = table.rows.filter((row) => row[index] === status).length;
        if (count)
          offers.push({
            id: `${table.id}:${status}`,
            title: `${count} ${status}`,
            count,
            detail: `Inspect ${status} rows${table.truncated ? ' in the visible result' : ''}`,
            artifactId: table.id,
            approval: 'none',
            action: { kind: 'filter', id: table.id, status },
          });
      }
    }
    return {
      artifacts,
      ...(this.selected ? { selected: this.selected } : {}),
      offers,
      ...(selected?.truncated
        ? { note: 'Result truncated. Refine the query before writing.' }
        : {}),
    };
  }
  dispose(): void {
    this.engine?.dispose();
    this.artifacts.clear();
    this.selected = undefined;
  }
}
