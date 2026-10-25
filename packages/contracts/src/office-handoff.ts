import { z } from 'zod';
import { SurfaceSchema } from './context.js';
import { NativeTargetSchema } from './host-profile.js';

const hash = z.string().regex(/^sha256:[a-f0-9]{64}$/);
/** Identity is supplied by the authenticated composition root, never by a document or model. */
export const OfficeApplicationBindingSchema = z
  .object({
    surface: SurfaceSchema,
    principalHash: hash,
    documentId: hash,
  })
  .strict();
export type OfficeApplicationBinding = z.infer<typeof OfficeApplicationBindingSchema>;

const metrics = z
  .object({
    template: z.string().min(1).max(16000),
    values: z
      .array(
        z
          .object({
            name: z.string().regex(/^[A-Za-z][A-Za-z0-9_]{0,63}$/),
            row: z.number().int().nonnegative(),
            column: z.number().int().nonnegative(),
          })
          .strict(),
      )
      .min(1)
      .max(100),
  })
  .strict();
const step = {
  id: z.string().regex(/^[a-z][a-z0-9-]{0,31}$/),
  application: z.string().regex(/^[a-z][a-z0-9-]{0,63}$/),
  binding: OfficeApplicationBindingSchema,
  metrics,
};
export const OfficeHandoffStepSchema = z.discriminatedUnion('kind', [
  z
    .object({
      ...step,
      kind: z.literal('managed-text'),
      capture: z.object({ target: NativeTargetSchema, text: z.string().max(64000) }).strict(),
      previousGeneratedText: z.string().max(64000).optional(),
    })
    .strict(),
  z
    .object({
      ...step,
      kind: z.literal('outlook-review'),
      subject: z.string().min(1).max(998),
    })
    .strict(),
]);
export type OfficeHandoffStep = z.infer<typeof OfficeHandoffStepSchema>;
export const OfficeHandoffPlanSchema = z
  .object({
    version: z.literal(1),
    id: z.string().regex(/^[a-z][a-z0-9-]{0,30}$/),
    sourceApplication: z.string().regex(/^[a-z][a-z0-9-]{0,63}$/),
    artifactId: z.string().regex(/^a_[a-f0-9]{24}$/),
    artifactHash: hash,
    sourceBinding: OfficeApplicationBindingSchema,
    steps: z.array(OfficeHandoffStepSchema).min(1).max(8),
  })
  .strict()
  .superRefine((plan, ctx) => {
    if (new Set(plan.steps.map((step) => step.id)).size !== plan.steps.length)
      ctx.addIssue({ code: z.ZodIssueCode.custom, message: 'Handoff step IDs must be unique.' });
    if (
      plan.steps.some(
        (step, index) => step.kind === 'outlook-review' && index !== plan.steps.length - 1,
      )
    )
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'A review form must be the final handoff step.',
      });
    if (new TextEncoder().encode(JSON.stringify(plan)).byteLength > 256 * 1024)
      ctx.addIssue({ code: z.ZodIssueCode.custom, message: 'Handoff plan exceeds 256 KiB.' });
  });
export type OfficeHandoffPlan = z.infer<typeof OfficeHandoffPlanSchema>;
