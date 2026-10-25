import { nativeRun } from './native-run.js';
import { documentOffset, type DocumentPage, type DocumentReadOptions } from '@ge/contracts';
import { isSet } from './capabilities-runtime.js';

const SHAPES_PER_PAGE = 40;
const TEXT_SHAPES = new Set(['TextBox', 'GeometricShape', 'Placeholder', 'Callout', 'Freeform']);

/** Load IDs first, then at most forty text frames in one sync. No per-shape sync loop. */
export async function readPowerPointPage(options: DocumentReadOptions): Promise<DocumentPage> {
  options.signal?.throwIfAborted();
  if (!isSet('PowerPointApi', '1.4'))
    throw new Error('PowerPointApi 1.4 is required for document reads.');
  const cursor = options.cursor?.split(':') ?? ['0', '0'];
  if (cursor.length !== 2) throw new Error('Invalid slide cursor.');
  const slideIndex = documentOffset(cursor[0]);
  const shapeOffset = documentOffset(cursor[1]);
  return nativeRun(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items/id');
    await ctx.sync();
    options.signal?.throwIfAborted();
    const slide = slides.items[slideIndex];
    const limitations = [
      'Text frames only: speaker notes, images, charts, tables, grouped shapes and other embedded objects are not extracted.',
    ];
    if (!slide) {
      if (slideIndex !== 0 || shapeOffset !== 0)
        throw new Error('The deck changed. Start a new read.');
      return { scope: 'PowerPoint slide text', units: [], limitations };
    }
    const shapes = slide.shapes;
    shapes.load({ select: 'id,type', skip: shapeOffset, top: SHAPES_PER_PAGE + 1 });
    await ctx.sync();
    if (shapes.items.length > SHAPES_PER_PAGE + 1)
      throw new Error('PowerPoint did not honor paged reads.');
    if (shapeOffset > 0 && shapes.items.length === 0)
      throw new Error('The slide changed. Start a new read.');
    const targets = shapes.items.slice(0, SHAPES_PER_PAGE);
    const text = targets
      .filter((shape) => TEXT_SHAPES.has(String(shape.type)))
      .map((shape) => ({ shape, range: shape.textFrame.textRange }));
    for (const item of text) item.range.load('text');
    await ctx.sync();
    options.signal?.throwIfAborted();
    const units: DocumentPage['units'] = targets.map((shape) => ({
      id: `slide:${slideIndex + 1}:${slide.id}/shape:${shape.id}`,
      parentId: `slide:${slideIndex + 1}:${slide.id}`,
      text:
        text.find((item) => item.shape.id === shape.id)?.range.text ??
        `[${shape.type}: content not extracted]`,
    }));
    if (shapeOffset === 0)
      units.unshift({
        id: `slide:${slideIndex + 1}:${slide.id}`,
        text: `Slide ${slideIndex + 1} of ${slides.items.length}`,
      });
    const nextCursor =
      shapes.items.length > SHAPES_PER_PAGE
        ? `${slideIndex}:${shapeOffset + SHAPES_PER_PAGE}`
        : slideIndex + 1 < slides.items.length
          ? `${slideIndex + 1}:0`
          : undefined;
    return {
      scope: 'PowerPoint slide text',
      units,
      limitations,
      ...(nextCursor ? { nextCursor } : {}),
    };
  });
}
