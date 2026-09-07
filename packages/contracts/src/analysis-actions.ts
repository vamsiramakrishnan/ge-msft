import { z } from 'zod';

/** Shared analysis command inputs. Execution remains in runtime; arithmetic remains in compute. */
export const ReconciliationSpecSchema = z
  .object({
    left: z.string(),
    right: z.string(),
    leftKey: z.number().int().nonnegative(),
    rightKey: z.number().int().nonnegative(),
    leftAmount: z.number().int().nonnegative(),
    rightAmount: z.number().int().nonnegative(),
    leftCurrency: z.number().int().nonnegative().optional(),
    rightCurrency: z.number().int().nonnegative().optional(),
    currency: z
      .string()
      .regex(/^[A-Z]{3}$/)
      .optional(),
    tolerance: z
      .string()
      .regex(/^\d{1,12}(\.\d{1,6})?$/)
      .default('0.01'),
  })
  .refine(
    (s) => Boolean(s.currency) || (s.leftCurrency !== undefined && s.rightCurrency !== undefined),
    'Choose a currency or a currency column in both tables.',
  );
export type ReconciliationSpec = z.infer<typeof ReconciliationSpecSchema>;

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

/** Artifact-producing actions accepted by the CLI, SDK and runtime binding environment. */
export const ANALYSIS_BINDING_KINDS = [
  'capture',
  'query',
  'reconcile',
  'filter',
  'inspect',
] as const;
export type AnalysisBindingKind = (typeof ANALYSIS_BINDING_KINDS)[number];
export function isAnalysisBindingKind(value: unknown): value is AnalysisBindingKind {
  return (ANALYSIS_BINDING_KINDS as readonly unknown[]).includes(value);
}
