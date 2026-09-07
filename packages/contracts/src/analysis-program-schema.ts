import { z } from 'zod';
import { AnalysisActionSchema } from './analysis-actions.js';

export const ANALYSIS_BINDING_NAME_PATTERN = /^[A-Za-z_][A-Za-z0-9_]*$/;

export const AnalysisProgramSchema = z
  .object({
    version: z.literal(1),
    steps: z
      .array(
        z.discriminatedUnion('op', [
          z.object({
            op: z.literal('bind'),
            name: z.string().max(64).regex(ANALYSIS_BINDING_NAME_PATTERN),
            action: AnalysisActionSchema,
          }),
          z.object({
            op: z.literal('materialize'),
            id: z.string().min(1),
            destination: z.string().min(1),
            whenNonEmpty: z.boolean().optional(),
          }),
        ]),
      )
      .min(1)
      .max(31),
    completion: z.literal('verified').default('verified'),
  })
  .strict();
export type AnalysisProgram = z.input<typeof AnalysisProgramSchema>;
