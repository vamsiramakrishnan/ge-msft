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

export const DeckSpecSchema = z.object({
  title: z.string().max(500).optional(),
  subject: z.string().max(1000).optional(),
  author: z.string().max(200).optional(),
  layout: z.enum(['wide', 'standard']).default('wide'),
  slides: z.array(DeckSlideSpecSchema).min(1).max(80),
});
export type DeckSpec = z.infer<typeof DeckSpecSchema>;
