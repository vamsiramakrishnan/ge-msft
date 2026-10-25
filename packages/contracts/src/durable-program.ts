import { z } from 'zod';
import { SurfaceSchema } from './context.js';
import { SourceVersionSchema } from './analysis.js';

/** An authored input declaration, never inferred from model output or restored as approval. */
export const DurableProgramEvidenceSchema = z.discriminatedUnion('kind', [
  z.object({ kind: z.literal('literal') }).strict(),
  z
    .object({
      kind: z.literal('cell-sources'),
      sources: z.array(SourceVersionSchema.strict()).min(1).max(16),
    })
    .strict(),
]);

/** Each checkpoint is a self-contained command program with one host effect. Bindings are stage-local. */
export const DurableProgramStageSchema = z
  .object({
    id: z.string().regex(/^[a-z][a-z0-9-]{0,63}$/),
    dependsOn: z
      .array(z.string().regex(/^[a-z][a-z0-9-]{0,63}$/))
      .max(31)
      .default([]),
    program: z
      .string()
      .min(1)
      .max(64 * 1024),
  })
  .strict();
export const DurableCommandProgramSchema = z
  .object({
    version: z.literal(1),
    id: z.string().regex(/^[a-z][a-z0-9-]{0,63}$/),
    surface: SurfaceSchema,
    stages: z.array(DurableProgramStageSchema).min(1).max(32),
    /** Missing legacy lineage is inspection-only for the built-in resume path. */
    evidence: DurableProgramEvidenceSchema.optional(),
  })
  .strict()
  .superRefine((program, ctx) => {
    const seen = new Set<string>();
    if (program.evidence?.kind === 'cell-sources') {
      const keys = program.evidence.sources.map(
        (source) => `${source.surface}:${source.documentId}:${source.locator}`,
      );
      if (new Set(keys).size !== keys.length)
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: 'Declared source versions must be unique.',
        });
    }
    for (const stage of program.stages) {
      if (seen.has(stage.id) || stage.dependsOn.some((id) => !seen.has(id)))
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: 'Stages need unique IDs and earlier dependencies.',
        });
      if (new Set(stage.dependsOn).size !== stage.dependsOn.length)
        ctx.addIssue({ code: z.ZodIssueCode.custom, message: 'Duplicate stage dependency.' });
      seen.add(stage.id);
    }
    if (new TextEncoder().encode(JSON.stringify(program)).byteLength > 256 * 1024)
      ctx.addIssue({ code: z.ZodIssueCode.custom, message: 'Durable program exceeds 256 KiB.' });
  });
export type DurableCommandProgram = z.infer<typeof DurableCommandProgramSchema>;
export type DurableProgramStage = z.infer<typeof DurableProgramStageSchema>;
