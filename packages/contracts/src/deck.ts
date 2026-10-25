import { z } from 'zod';

const HexColorSchema = z.string().regex(/^[0-9A-Fa-f]{6}$/);
const SlideUnitSchema = z.number().finite().min(0).max(20);

export const DeckTextStyleSchema = z.object({
  fontFace: z.string().min(1).max(80).optional(),
  fontSize: z.number().finite().min(6).max(72).optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  color: HexColorSchema.optional(),
  align: z.enum(['left', 'center', 'right']).optional(),
});
export type DeckTextStyle = z.infer<typeof DeckTextStyleSchema>;

const DeckElementFrameSchema = z.object({
  x: SlideUnitSchema,
  y: SlideUnitSchema,
  w: SlideUnitSchema,
  h: SlideUnitSchema,
  allowOverlap: z.boolean().optional(),
});

export const DeckElementSchema = z.discriminatedUnion('kind', [
  DeckElementFrameSchema.extend({
    kind: z.literal('text'),
    text: z.string().min(1).max(4000),
    style: DeckTextStyleSchema.optional(),
  }),
  DeckElementFrameSchema.extend({
    kind: z.literal('bullets'),
    items: z.array(z.string().min(1).max(500)).min(1).max(30),
    style: DeckTextStyleSchema.optional(),
  }),
  DeckElementFrameSchema.extend({
    kind: z.literal('table'),
    rows: z
      .array(z.array(z.string().max(500)).min(1).max(12))
      .min(1)
      .max(40),
    style: DeckTextStyleSchema.optional(),
  }),
  DeckElementFrameSchema.extend({
    kind: z.literal('chart'),
    chartType: z.enum(['column', 'bar', 'line', 'pie', 'area']),
    labels: z.array(z.string().max(200)).min(1).max(500),
    series: z
      .array(
        z
          .object({
            name: z.string().min(1).max(200),
            values: z.array(z.number().finite()).min(1).max(500),
          })
          .strict(),
      )
      .min(1)
      .max(12),
    title: z.string().max(500).optional(),
    showLegend: z.boolean().default(true),
  }),
  DeckElementFrameSchema.extend({
    kind: z.literal('image'),
    mimeType: z.enum(['image/png', 'image/jpeg']),
    base64: z.string().min(1).max(8_000_000),
    altText: z.string().max(500).optional(),
  }),
]);
export type DeckElement = z.infer<typeof DeckElementSchema>;

export const DeckSlideSpecSchema = z.object({
  id: z.string().min(1).max(120).optional(),
  title: z.string().max(500).optional(),
  subtitle: z.string().max(1000).optional(),
  notes: z.string().max(8000).optional(),
  backgroundColor: HexColorSchema.optional(),
  elements: z.array(DeckElementSchema).max(50).default([]),
});
export type DeckSlideSpec = z.infer<typeof DeckSlideSpecSchema>;

export const DeckSpecSchema = z
  .object({
    title: z.string().max(500).optional(),
    subject: z.string().max(1000).optional(),
    author: z.string().max(200).optional(),
    layout: z.enum(['wide', 'standard']).default('wide'),
    theme: z
      .object({
        headFontFace: z.string().min(1).max(80).optional(),
        bodyFontFace: z.string().min(1).max(80).optional(),
        headingColor: HexColorSchema.optional(),
        bodyColor: HexColorSchema.optional(),
        accentColors: z.array(HexColorSchema).min(1).max(12).optional(),
      })
      .strict()
      .optional(),
    slides: z.array(DeckSlideSpecSchema).min(1).max(80),
  })
  .superRefine((spec, context) => {
    let bytes = 0;
    for (const [slideIndex, slide] of spec.slides.entries()) {
      bytes +=
        (slide.title?.length ?? 0) * 4 +
        (slide.subtitle?.length ?? 0) * 4 +
        (slide.notes?.length ?? 0) * 4;
      for (const [elementIndex, element] of slide.elements.entries()) {
        if (element.kind === 'image') bytes += element.base64.length;
        else if (element.kind === 'chart') {
          bytes +=
            element.labels.join('').length * 4 +
            element.series.reduce(
              (sum, series) => sum + series.values.length * 24 + series.name.length * 4,
              0,
            );
          if (element.series.some((series) => series.values.length !== element.labels.length))
            context.addIssue({
              code: z.ZodIssueCode.custom,
              path: ['slides', slideIndex, 'elements', elementIndex],
              message: 'Every chart series must match the label count.',
            });
          if (element.chartType === 'pie' && element.series.length !== 1)
            context.addIssue({
              code: z.ZodIssueCode.custom,
              path: ['slides', slideIndex, 'elements', elementIndex],
              message: 'A pie chart requires exactly one series.',
            });
        } else bytes += JSON.stringify(element).length * 4;
      }
    }
    if (bytes > 16 * 1024 * 1024)
      context.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'Deck content exceeds the aggregate 16 MiB budget.',
      });
  });
export type DeckSpec = z.infer<typeof DeckSpecSchema>;
