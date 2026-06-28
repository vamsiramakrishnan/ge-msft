import type {
  ActuationKind,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { buildDocStateSnapshot } from '@ge/content';
import { POWERPOINT_CAPABILITIES } from './capabilities.js';
import { isSet } from './capabilities-runtime.js';
import {
  parseSlideSelector,
  searchSlides,
  selectedShapeToContext,
  selectedSlideToContext,
  shapeContextRef,
  slideContextRef,
  shapesToSlideText,
  slideElementsToDocStateBlocks,
  slidesToContext,
  type SlideElement,
} from './capture.js';
import { planInsertSlide } from './actuate-plan.js';
import { documentChanged, selectionChanged } from './events.js';

/**
 * Upper bound on slides materialized by a single read port (`captureDocState`/`searchDocument`/
 * `readRange`). A deck can be large; reading every shape's text across the Office.js bridge is
 * O(slides × shapes) syncs, so we cap the scan to keep a per-turn read bounded (ADR-0006 — every
 * read is bounded). A deck over this is read as its first {@link MAX_READ_SLIDES} slides.
 */
export const MAX_READ_SLIDES = 60;

/**
 * The PowerPoint `DocBridge`. The ONLY place Office.js (`PowerPoint.run`) is touched. Reads via
 * the native object model (selected slides → shapes' text + speaker notes); writes by composing
 * slides into the deck (`insertSlidesFromBase64` for a prebuilt deck, else `slides.add()` +
 * placeholder text) and setting speaker notes — each reversible and provenanced via the shared
 * `ActuationResult` shape. Pure mapping lives in `capture.ts` / `actuate-plan.ts` (unit-tested);
 * this file is the host wiring.
 *
 * Requirement-set versions used (confirmed against node_modules/@types/office-js/index.d.ts):
 *   - `Presentation.getSelectedSlides()` / `getSelectedShapes()` → PowerPointApi 1.5 (l.178785 / l.178776).
 *   - `Slide.shapes` (ShapeCollection) → PowerPointApi 1.3 (l.186098); `Shape.textFrame` → 1.4 (l.186557);
 *     `TextRange.text` (read/write) → PowerPointApi 1.4 (l.180262-180267).
 *   - `Slide.id` → 1.2 (l.186126), `Slide.index` → 1.8 (l.186133).
 *   - `SlideCollection.add()` → PowerPointApi 1.3 (l.187424); `insertSlidesFromBase64` → 1.2 (l.178812).
 *   - PowerPoint exposes NO object-model selection/change event in this typings, so `watch` uses the
 *     Office-level `Office.EventType.DocumentSelectionChanged` (l.645) + `ActiveViewChanged` (l.582)
 *     with add/removeHandlerAsync (l.3875 / l.3965). Neither carries a coauthor source → origin 'local'.
 */
/**
 * The exact `ActuationKind`s {@link PowerPointBridge.actuate} handles (ADR-0006 closure source of
 * truth). `set-speaker-notes` is deliberately ABSENT — it had no working write path (always
 * degraded) and was un-advertised. The conformance test asserts this equals the manifest's kinds.
 */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = ['insert-slide', 'set-shape-text'];

export class PowerPointBridge implements DocBridge {
  readonly surface = 'powerpoint' as const;

  getCapabilities(): CapabilityManifest {
    return POWERPOINT_CAPABILITIES;
  }

  async listContext(): Promise<ContextRef[]> {
    // getSelectedSlides is PowerPointApi 1.5; on an older host we can still list the deck.
    if (!isSet('PowerPointApi', '1.5')) {
      return [{ id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' }];
    }
    return PowerPoint.run(async (ctx) => {
      const selected = ctx.presentation.getSelectedSlides();
      selected.load('items/id,items/index');
      await ctx.sync();

      const refs: ContextRef[] = [];
      const first = selected.items[0];
      if (first) {
        const slide = await readSlide(ctx, first);
        refs.push(slideContextRef(slide));
        for (const [index, shape] of (slide.shapes ?? []).entries()) {
          refs.push(shapeContextRef(slide, shape, index));
        }
      }
      refs.push({ id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' });
      return refs;
    });
  }

  async resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    if (ref.kind === 'shape') {
      const target = shapeRevealTarget(ref);
      if (!target) return [];
      return PowerPoint.run(async (ctx) => readShapeContext(ctx, target));
    }
    if (ref.kind === 'slide') {
      const target = slideRevealTarget(ref);
      if (!target) return [];
      try {
        return await PowerPoint.run(async (ctx) => {
          const slide = ctx.presentation.slides.getItem(target.slideId);
          const element = await readSlide(ctx, slide);
          return selectedSlideToContext(element);
        });
      } catch {
        return [];
      }
    }
    if (ref.kind === 'selection') {
      return PowerPoint.run(async (ctx) => {
        const selected = ctx.presentation.getSelectedSlides();
        selected.load('items/id,items/index');
        await ctx.sync();
        const slide = selected.items[0];
        if (!slide) return [];
        const element = await readSlide(ctx, slide);
        return selectedSlideToContext(element);
      });
    }
    // Whole deck → each slide's shapes + (best-effort) speaker notes → native blocks → chunks.
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id,items/index');
      await ctx.sync();
      const elements: SlideElement[] = [];
      for (const slide of slides.items) {
        elements.push(await readSlide(ctx, slide));
      }
      return slidesToContext('pp:deck', 'Whole deck', elements);
    });
  }

  /** Monotonic `<doc_state>` version, bumped on each capture (ADR-0003 Layer B element 1). */
  private docStateVersion = 0;

  /**
   * ADR-0003 Layer B element 1 / ADR-0006 `outline` read: an ambient structural snapshot of the
   * deck — a slide inventory (per-slide title + body) read from the native object model and mapped
   * through the same `native.slide()` blocks grounding context uses. Bounded to
   * {@link MAX_READ_SLIDES} slides; the snapshot's `inventory` lists each slide. Reading the deck
   * needs `TextRange.text` (PowerPointApi 1.4) — on an older host we yield `undefined` (the runtime
   * just streams without the ambient part). Empty deck → `undefined`. Version increments per capture.
   */
  async captureDocState(): Promise<DocStateSnapshot | undefined> {
    if (!isSet('PowerPointApi', '1.4')) return undefined;
    const slides = await this.readAllSlides();
    if (slides.length === 0) return undefined;
    this.docStateVersion += 1;
    return buildDocStateSnapshot({
      surface: 'powerpoint',
      version: this.docStateVersion,
      blocks: slideElementsToDocStateBlocks(slides),
    });
  }

  /**
   * ADR-0006 `search` read: scan the deck's slide text for `query` and return matching slides as
   * `ResolvedContext` data (never instructions), bounded by `searchSlides`. Reads via the native
   * model (gated on PowerPointApi 1.4); empty query / older host / no match → `[]`.
   */
  async searchDocument(query: string): Promise<ResolvedContext[]> {
    const q = query.trim();
    if (!q || !isSet('PowerPointApi', '1.4')) return [];
    const slides = await this.readAllSlides();
    return searchSlides(slides, q);
  }

  /**
   * ADR-0006 addressable `read <slide:N>` verb: resolve a slide selector (`slide:N` / `slide N` /
   * bare 1-based `N`) to that single slide's text as `ResolvedContext` data. Unaddressable selectors
   * (a name, junk) / out-of-range index / older host / empty deck → `[]` — the bridge degrades rather
   * than guessing. Reads via the native model (gated on PowerPointApi 1.4); the read is one slide, so
   * it is inherently bounded.
   */
  async readRange(selector: string): Promise<ResolvedContext[]> {
    const index = parseSlideSelector(selector);
    if (index === undefined || !isSet('PowerPointApi', '1.4')) return [];
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id,items/index');
      await ctx.sync();
      const slide = slides.items[index];
      if (!slide) return [];
      const element = await readSlide(ctx, slide);
      return selectedSlideToContext(element);
    });
  }

  canRevealContext(ref: ContextRef): boolean {
    return ref.surface === 'powerpoint' && powerpointRevealTarget(ref) !== undefined;
  }

  async revealContext(ref: ContextRef): Promise<void> {
    const target = powerpointRevealTarget(ref);
    if (!target) return;
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.setSelectedSlides([target.slideId]);
      if (target.shapeId) {
        const slide = ctx.presentation.slides.getItem(target.slideId);
        slide.setSelectedShapes([target.shapeId]);
      }
      await ctx.sync();
    });
  }

  /**
   * Read up to {@link MAX_READ_SLIDES} slides of the deck into pure {@link SlideElement}s — the
   * shared host read behind `captureDocState`/`searchDocument`. Bounded so a huge deck can't blow
   * the per-turn budget; read-only (loads shape text, writes nothing).
   */
  private async readAllSlides(): Promise<SlideElement[]> {
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id,items/index');
      await ctx.sync();
      const elements: SlideElement[] = [];
      for (const slide of slides.items.slice(0, MAX_READ_SLIDES)) {
        elements.push(await readSlide(ctx, slide));
      }
      return elements;
    });
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    switch (req.kind) {
      case 'insert-slide':
        return this.applyInsertSlide(req);
      case 'set-shape-text':
        return this.applySetShapeText(req);
      default:
        return {
          ok: false,
          changeId: req.changeId,
          kind: req.kind,
          error: { code: 'unsupported', message: `PowerPoint bridge cannot ${req.kind}` },
        };
    }
  }

  private async applySetShapeText(req: ActuationRequest): Promise<ActuationResult> {
    const slideId = req.params.target?.slideId;
    const shapeId = req.params.target?.shapeId;
    const text = req.params.text;
    if (!slideId || !shapeId) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'no_target',
          message: 'set-shape-text needs target.slideId and target.shapeId',
        },
      };
    }
    if (text === undefined) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_text', message: 'set-shape-text needs params.text' },
      };
    }
    if (!isSet('PowerPointApi', '1.4')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'unsupported', message: 'PowerPointApi 1.4 is required for shape text.' },
      };
    }

    try {
      return await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.slides.getItem(slideId);
        const shape = slide.shapes.getItemOrNullObject(shapeId);
        shape.load('isNullObject');
        await ctx.sync();
        if (shape.isNullObject) {
          return {
            ok: false,
            changeId: req.changeId,
            kind: req.kind,
            degraded: true,
            error: {
              code: 'target_conflict',
              message: 'The selected PowerPoint shape no longer exists.',
            },
          };
        }
        const range = shape.textFrame.textRange;
        range.load('text');
        await ctx.sync();
        const priorText = range.text ?? '';
        range.text = text;
        await ctx.sync();
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${slideId}:${shapeId}`,
          inverse: { op: 'restore-text', anchor: `pp:shape:${slideId}:${shapeId}`, priorText },
        };
      });
    } catch {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint shape could not be re-read before writing.',
        },
      };
    }
  }

  private async applyInsertSlide(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planInsertSlide(req);
    if (!plan.base64 && !plan.title && plan.bullets.length === 0) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'empty_slide', message: 'insert-slide needs params.slide or params.ooxml' },
      };
    }
    return PowerPoint.run(async (ctx) => {
      // Prebuilt deck path: the agent supplied a Base64 PPTX — let the host merge it (1.2).
      if (plan.base64) {
        const options = await deckInsertOptions(ctx, plan);
        ctx.presentation.insertSlidesFromBase64(plan.base64, options);
        await ctx.sync();
        const location =
          plan.slideCount === undefined ? 'inserted-deck' : `inserted-deck:${plan.slideCount}`;
        return { ok: true, changeId: req.changeId, kind: req.kind, location };
      }
      // Native compose path: append a slide, then fill its placeholder shapes with title/body.
      const slides = ctx.presentation.slides;
      const before = slides.getCount();
      await ctx.sync();
      slides.add();
      await ctx.sync();
      // Re-read the appended slide (last index) and write its shapes' text.
      const slide = slides.getItemAt(before.value);
      const shapes = slide.shapes;
      shapes.load('items/id');
      slide.load('id');
      await ctx.sync();
      writeSlideText(shapes, plan.title, plan.bullets);
      await ctx.sync();
      return {
        ok: true,
        changeId: req.changeId,
        kind: req.kind,
        location: `slide:${slide.id}`,
      };
    });
  }

  // NOTE: a `set-speaker-notes` actuation was handled here but ALWAYS degraded — this Office.js
  // typings version exposes no `Slide.notes`/notesSlide write path. Per ADR-0006 we removed the
  // phantom from the manifest AND its `actuate()` case (advertised==handled), rather than keep a
  // case that can never succeed. Re-add (manifest + case + CLI verb) once the host typings ship a
  // notes writer. The pure `planSpeakerNotes` plan stays in `actuate-plan.ts` for that day.

  /**
   * Stream PowerPoint host events into the trigger engine via the Office-level event bus
   * (PowerPoint has no object-model selection/change event in this typings). Each registration
   * is defensive: a failed/absent registration simply means we never emit that event — it never
   * throws. Returns an `Unsubscribe` that removes every handler we added.
   */
  watch(emit: (event: HostEvent) => void): Unsubscribe {
    let onSelection: (() => void) | undefined;
    let onView: (() => void) | undefined;

    try {
      const handler = (): void => emit(selectionChanged());
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
      onSelection = handler;
    } catch {
      // Selection observation unavailable on this host — simply don't emit it.
    }

    try {
      const handler = (): void => emit(documentChanged());
      Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, handler);
      onView = handler;
    } catch {
      // Slide-navigation observation unavailable — skip.
    }

    let unsubscribed = false;
    return () => {
      if (unsubscribed) return;
      unsubscribed = true;
      if (onSelection) {
        try {
          Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
            handler: onSelection,
          });
        } catch {
          // best-effort teardown
        } finally {
          onSelection = undefined;
        }
      }
      if (onView) {
        try {
          Office.context.document.removeHandlerAsync(Office.EventType.ActiveViewChanged, {
            handler: onView,
          });
        } catch {
          // best-effort teardown
        } finally {
          onView = undefined;
        }
      }
    };
  }
}

async function deckInsertOptions(
  ctx: PowerPoint.RequestContext,
  plan: ReturnType<typeof planInsertSlide>,
): Promise<PowerPoint.InsertSlideOptions | undefined> {
  const options: PowerPoint.InsertSlideOptions = {};
  if (plan.formatting !== undefined) options.formatting = plan.formatting;
  const targetSlideId = plan.targetSlideId ?? (await appendTargetSlideId(ctx, plan.targetIndex));
  if (targetSlideId !== undefined) options.targetSlideId = targetSlideId;
  return Object.keys(options).length === 0 ? undefined : options;
}

async function appendTargetSlideId(
  ctx: PowerPoint.RequestContext,
  targetIndex: number | undefined,
): Promise<string | undefined> {
  const slides = ctx.presentation.slides;
  const count = slides.getCount();
  await ctx.sync();
  if (count.value <= 0) return undefined;
  const index = targetIndex ?? count.value - 1;
  const slide = slides.getItemAt(index);
  slide.load('id');
  await ctx.sync();
  return slide.id;
}

interface PowerPointRevealTarget {
  slideId: string;
  shapeId?: string;
}

function powerpointRevealTarget(ref: ContextRef): PowerPointRevealTarget | undefined {
  if (ref.surface !== 'powerpoint' || !isSet('PowerPointApi', '1.5')) return undefined;
  return shapeRevealTarget(ref) ?? slideRevealTarget(ref);
}

function slideRevealTarget(ref: ContextRef): PowerPointRevealTarget | undefined {
  const id =
    prefixedValue(ref.id, 'pp:slide:', 'slide:') ?? prefixedValue(ref.anchor?.locator, 'slide:');
  return id ? { slideId: id } : undefined;
}

function shapeRevealTarget(ref: ContextRef): PowerPointRevealTarget | undefined {
  const fromId = prefixedValue(ref.id, 'pp:shape:', 'shape:');
  const fromLocator = prefixedValue(ref.anchor?.locator, 'pp:shape:', 'shape:');
  const raw = fromId ?? fromLocator;
  if (!raw) return undefined;
  const [slideId, shapeId] = raw.split(':');
  if (!slideId || !shapeId) return undefined;
  return { slideId, shapeId };
}

function prefixedValue(value: string | undefined, ...prefixes: string[]): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) return undefined;
  for (const prefix of prefixes) {
    if (trimmed.startsWith(prefix)) {
      const rest = trimmed.slice(prefix.length).trim();
      return rest || undefined;
    }
  }
  return undefined;
}

/** Read one slide's shapes' text (+ id/index) into a pure {@link SlideElement}. */
async function readSlide(
  ctx: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
): Promise<SlideElement> {
  const shapes = slide.shapes;
  shapes.load('items/id');
  slide.load('id,index');
  await ctx.sync();

  const ranges: PowerPoint.TextRange[] = [];
  for (const shape of shapes.items) {
    const range = shape.textFrame.textRange;
    range.load('text');
    ranges.push(range);
  }
  await ctx.sync();

  const shapeTexts = ranges.map((r) => r.text ?? '');
  const { title, body } = shapesToSlideText(shapeTexts);
  return {
    index: slide.index,
    slideId: slide.id,
    title,
    body,
    shapes: shapes.items.map((shape, index) => ({
      shapeId: shape.id,
      text: shapeTexts[index] ?? '',
    })),
  };
}

async function readShapeContext(
  ctx: PowerPoint.RequestContext,
  target: PowerPointRevealTarget,
): Promise<ResolvedContext[]> {
  const slide = ctx.presentation.slides.getItem(target.slideId);
  slide.load('id,index');
  const shape = slide.shapes.getItemOrNullObject(target.shapeId ?? '');
  shape.load('id,isNullObject');
  await ctx.sync();
  if (shape.isNullObject || !target.shapeId) return [];
  const range = shape.textFrame.textRange;
  range.load('text');
  await ctx.sync();
  return selectedShapeToContext(
    { index: slide.index, slideId: slide.id },
    { shapeId: shape.id, text: range.text ?? '' },
  );
}

/**
 * Write a composed title + bullets into a freshly added slide. The first shape (the title
 * placeholder, by layout convention) takes the title; the next gets the bullets joined by
 * newlines. Shapes are addressed by their loaded items; missing placeholders are skipped.
 */
function writeSlideText(
  shapes: PowerPoint.ShapeCollection,
  title: string,
  bullets: string[],
): void {
  const titleShape = shapes.items[0];
  if (titleShape && title) titleShape.textFrame.textRange.text = title;
  const bodyShape = shapes.items[1];
  if (bodyShape && bullets.length > 0) {
    bodyShape.textFrame.textRange.text = bullets.join('\n');
  }
}
