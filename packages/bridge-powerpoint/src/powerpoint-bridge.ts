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
  selectedSlideToContext,
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
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = ['insert-slide'];

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
        refs.push({
          id: `pp:slide:${first.id}`,
          kind: 'slide',
          surface: 'powerpoint',
          title: `Slide ${first.index + 1}`,
          live: true,
        });
      }
      refs.push({ id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' });
      return refs;
    });
  }

  async resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    if (ref.kind === 'slide' || ref.kind === 'selection' || ref.kind === 'shape') {
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
      default:
        return {
          ok: false,
          changeId: req.changeId,
          kind: req.kind,
          error: { code: 'unsupported', message: `PowerPoint bridge cannot ${req.kind}` },
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
        ctx.presentation.insertSlidesFromBase64(plan.base64);
        await ctx.sync();
        return { ok: true, changeId: req.changeId, kind: req.kind, location: 'inserted-deck' };
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
  return { index: slide.index, slideId: slide.id, title, body };
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
