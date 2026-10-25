import { contentHash, DeckTemplateCaptureSchema, type DeckTemplateCapture } from '@ge/contracts';
import { nativeRun } from './native-run.js';
import { isSet } from './capabilities-runtime.js';

type Layout = DeckTemplateCapture['masters'][number]['layouts'][number];
/** Geometry, identity and placeholder roles only; never claims to clone unsupported master assets. */
export async function readNativeLayout(
  ctx: PowerPoint.RequestContext,
  layout: PowerPoint.SlideLayout,
  shapeLimit: number,
): Promise<{ layout: Layout; omitted: number; loaded: number }> {
  layout.load(isSet('PowerPointApi', '1.8') ? 'id,name,type' : 'id,name');
  const count = layout.shapes.getCount();
  await ctx.sync();
  const shapes = Array.from({ length: Math.min(count.value, shapeLimit) }, (_, index) => {
    const shape = layout.shapes.getItemAt(index);
    shape.load('id,type,left,top,width,height');
    return shape;
  });
  await ctx.sync();
  const placeholders = shapes.filter((shape) => shape.type === 'Placeholder');
  if (isSet('PowerPointApi', '1.8')) {
    for (const shape of placeholders) shape.placeholderFormat.load('type');
    await ctx.sync();
  }
  const value = {
    id: layout.id,
    name: layout.name.slice(0, 256),
    ...(isSet('PowerPointApi', '1.8') ? { type: String(layout.type) } : {}),
    placeholders: placeholders.map((shape) => ({
      id: shape.id,
      type: isSet('PowerPointApi', '1.8') ? String(shape.placeholderFormat.type) : 'unknown',
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
    })),
  };
  return {
    layout: {
      ...value,
      revision: await contentHash({
        ...value,
        shapes: shapes.map((shape) => ({
          id: shape.id,
          type: shape.type,
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height,
        })),
      }),
    },
    omitted: count.value - shapes.length,
    loaded: shapes.length,
  };
}

export async function captureNativeDeckTemplate(documentKey: string): Promise<DeckTemplateCapture> {
  if (!isSet('PowerPointApi', '1.3'))
    throw new Error('PowerPointApi 1.3 is required for native template capture.');
  return nativeRun(async (ctx) => {
    const source = ctx.presentation.slideMasters;
    const count = source.getCount();
    await ctx.sync();
    const omissions: string[] = [];
    if (count.value > 16)
      omissions.push(`${count.value - 16} additional masters were not captured.`);
    const masters: DeckTemplateCapture['masters'] = [];
    let remainingLayouts = 128;
    let remainingShapes = 256;
    for (let index = 0; index < Math.min(count.value, 16); index++) {
      const master = source.getItemAt(index);
      master.load('id,name');
      const layoutCount = master.layouts.getCount();
      await ctx.sync();
      const capturedCount = Math.min(layoutCount.value, 32, remainingLayouts);
      if (capturedCount < layoutCount.value)
        omissions.push(
          `Master ${index + 1}: ${layoutCount.value - capturedCount} layouts omitted.`,
        );
      const layouts: Layout[] = [];
      for (let layoutIndex = 0; layoutIndex < capturedCount; layoutIndex++) {
        const captured = await readNativeLayout(
          ctx,
          master.layouts.getItemAt(layoutIndex),
          Math.min(128, remainingShapes),
        );
        layouts.push(captured.layout);
        // Budget all loaded shape metadata, including non-placeholder shapes.
        remainingShapes -= captured.loaded;
        if (captured.omitted)
          omissions.push(
            `Master ${index + 1}, layout ${layoutIndex + 1}: ${captured.omitted} shape metadata records omitted.`,
          );
      }
      remainingLayouts -= capturedCount;
      masters.push({ id: master.id, name: master.name.slice(0, 256), layouts });
    }
    return DeckTemplateCaptureSchema.parse({
      version: 1,
      documentId: await contentHash({ surface: 'powerpoint', documentKey }),
      complete: omissions.length === 0,
      omissions: omissions.slice(0, 32),
      masters,
    });
  });
}
