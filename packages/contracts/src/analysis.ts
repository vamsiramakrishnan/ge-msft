import { z } from 'zod';
import { SurfaceSchema } from './context.js';

export const MAX_SNAPSHOT_CELLS = 100_000;
export const CellValueSchema = z.union([
  z.string().max(32_768),
  z.number().finite(),
  z.boolean(),
  z.null(),
]);
export type CellValue = z.infer<typeof CellValueSchema>;
export const CellGridSchema = z
  .array(z.array(CellValueSchema).max(256))
  .min(1)
  .max(MAX_SNAPSHOT_CELLS)
  .refine(
    (rows) =>
      rows.length > 0 &&
      rows[0]!.length > 0 &&
      rows.every((r) => r.length === rows[0]!.length) &&
      rows.length * rows[0]!.length <= MAX_SNAPSHOT_CELLS,
    'A bounded rectangular grid is required',
  );
export const SourceVersionSchema = z.object({
  surface: SurfaceSchema,
  documentId: z.string().min(1).max(256),
  locator: z.string().min(1).max(1024),
  objectId: z.string().min(1).max(256).optional(),
  hash: z.string().regex(/^sha256:[a-f0-9]{64}$/),
});
export type SourceVersion = z.infer<typeof SourceVersionSchema>;
export const CellSnapshotSchema = SourceVersionSchema.extend({
  capturedAt: z.string().datetime(),
  values: CellGridSchema,
  formulas: z.array(z.array(z.string().max(32_768))).optional(),
}).refine(
  (s) =>
    !s.formulas ||
    (s.values.length > 0 &&
      s.formulas.length === s.values.length &&
      s.formulas.every((r) => r.length === s.values[0]!.length)),
  'Formula dimensions must match values',
);
export type CellSnapshot = z.infer<typeof CellSnapshotSchema>;
export const TableColumnSchema = z.object({
  name: z.string().regex(/^c\d+$/),
  label: z.string().max(512),
  type: z.enum(['string', 'number', 'boolean', 'mixed', 'empty']),
});
export const TableArtifactSchema = z
  .object({
    id: z.string().regex(/^a_[a-f0-9]{24}$/),
    title: z.string().min(1).max(256),
    hash: z.string().regex(/^sha256:[a-f0-9]{64}$/),
    createdAt: z.string().datetime(),
    columns: z.array(TableColumnSchema).min(1).max(256),
    rows: z.array(z.array(CellValueSchema)).max(MAX_SNAPSHOT_CELLS),
    sources: z.array(SourceVersionSchema).max(16),
    lineage: z.object({
      parents: z.array(z.string()).max(16),
      operation: z.enum(['snapshot', 'query', 'reconcile', 'profile']),
      expression: z.string().max(32_768).optional(),
    }),
    truncated: z.boolean().default(false),
  })
  .refine(
    (a) =>
      a.columns.every((c, i) => c.name === `c${i}`) &&
      a.rows.every((r) => r.length === a.columns.length) &&
      a.rows.length * a.columns.length <= MAX_SNAPSHOT_CELLS,
    'Artifact dimensions exceed the budget',
  );
export type TableArtifact = z.infer<typeof TableArtifactSchema>;
export type ArtifactSummary = Omit<TableArtifact, 'rows'> & {
  rowCount: number;
  preview: CellValue[][];
};
export const VerificationSchema = z.object({
  status: z.enum(['verified', 'mismatch', 'unknown']),
  beforeHash: z.string().optional(),
  afterHash: z.string().optional(),
  message: z.string().max(1000).optional(),
});
export type Verification = z.infer<typeof VerificationSchema>;

/** Hash content and addresses, never capture timestamps. The hash is a freshness token, not authority. */
export async function contentHash(value: unknown): Promise<string> {
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return `sha256:${Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, '0')).join('')}`;
}
export async function makeCellSnapshot(
  input: Omit<CellSnapshot, 'hash' | 'capturedAt'>,
): Promise<CellSnapshot> {
  const hash = await contentHash({
    surface: input.surface,
    documentId: input.documentId,
    locator: input.locator,
    objectId: input.objectId,
    values: input.values,
    formulas: input.formulas ?? [],
  });
  return CellSnapshotSchema.parse({ ...input, hash, capturedAt: new Date().toISOString() });
}
export function sourceVersion(snapshot: SourceVersion): SourceVersion {
  return SourceVersionSchema.parse(snapshot);
}
