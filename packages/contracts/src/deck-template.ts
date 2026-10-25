import { z } from 'zod';

/** Native metadata only. A captured layout is evidence, never permission to edit a slide. */
export const DeckTemplateCaptureSchema = z
  .object({
    version: z.literal(1),
    documentId: z.string().min(1).max(256),
    complete: z.boolean(),
    omissions: z.array(z.string().max(240)).max(32),
    masters: z
      .array(
        z
          .object({
            id: z.string().min(1).max(256),
            name: z.string().max(256),
            layouts: z
              .array(
                z
                  .object({
                    id: z.string().min(1).max(256),
                    name: z.string().max(256),
                    type: z.string().max(80).optional(),
                    revision: z.string().min(1).max(256),
                    placeholders: z
                      .array(
                        z
                          .object({
                            id: z.string().max(256),
                            type: z.string().max(80),
                            left: z.number().finite(),
                            top: z.number().finite(),
                            width: z.number().finite().nonnegative(),
                            height: z.number().finite().nonnegative(),
                          })
                          .strict(),
                      )
                      .max(128),
                  })
                  .strict(),
              )
              .max(32),
          })
          .strict(),
      )
      .max(16),
  })
  .strict()
  .superRefine((value, ctx) => {
    if (new TextEncoder().encode(JSON.stringify(value)).byteLength > 128 * 1024)
      ctx.addIssue({ code: z.ZodIssueCode.custom, message: 'Native template exceeds 128 KiB.' });
  });
export type DeckTemplateCapture = z.infer<typeof DeckTemplateCaptureSchema>;
