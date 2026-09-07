import { unknownActuationResult } from '@ge/contracts';
import { createBridgeDispatch } from '@ge/runtime';
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
 *   - `ShapeCollection.addTextBox` / `addGeometricShape` / `addLine` → PowerPointApi 1.4
 *     (l.184213 / l.184146 / l.184178) with `ShapeAddOptions` geometry (l.181900).
 *   - `Shape.fill` → 1.4 (l.186512; `ShapeFill.foregroundColor` l.182363), `Shape.lineFormat` → 1.4
 *     (l.186527; `ShapeLineFormat.color` l.186392), `TextRange.font` (`ShapeFont`) → 1.4
 *     (l.180230, class l.179862); `Shape.setZOrder` → PowerPointApi 1.8 (l.186782).
 *   - `ShapeCollection.addTable(rowCount, columnCount, TableAddOptions)` → PowerPointApi 1.8
 *     (l.184202; options l.184011); `Shape.getTable()` → 1.8 (l.186746); `Table.getCellOrNullObject`
 *     → 1.8 (l.183654); `TableCell.text` (read/write) → 1.8 (l.182631).
 *   - There is NO PowerPoint image-insertion API in this typings version (`addImage` exists only on
 *     Excel's ShapeCollection, l.55471; PPT's `getImageAsBase64` is read-only), so `insert-image`
 *     stays un-advertised for this surface.
 *   - PowerPoint exposes NO object-model selection/change event in this typings, so `watch` uses the
 *     Office-level `Office.EventType.DocumentSelectionChanged` (l.645) + `ActiveViewChanged` (l.582)
 *     with add/removeHandlerAsync (l.3875 / l.3965). Neither carries a coauthor source → origin 'local'.
 */

export class PowerPointBridge implements DocBridge {
  private static readonly dispatcher = createBridgeDispatch<PowerPointBridge>(
    'powerpoint',
    {
      'insert-slide': (host, request) => host.applyInsertSlide(request),
      'set-shape-text': (host, request) => host.applySetShapeText(request),
      'add-shape': (host, request) => host.applyAddShape(request),
      'format-shape': (host, request) => host.applyFormatShape(request),
      'add-table-slide': (host, request) => host.applyAddTableSlide(request),
    },
    { provenance: 'unsupported' },
  );
  static readonly handledActuations = PowerPointBridge.dispatcher.handledActuations;

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
    return PowerPointBridge.dispatcher.dispatch(this, req);
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

    let mutationQueued = false;
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
        mutationQueued = true;
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
      if (mutationQueued)
        return unknownActuationResult(
          req,
          'PowerPoint did not confirm the dispatched change. Inspect the slide before trying again.',
        );
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

  private async applyAddShape(req: ActuationRequest): Promise<ActuationResult> {
    const resolution = resolveAddShape(req);
    if (!resolution.ok) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: resolution.code, message: resolution.message },
      };
    }
    if (!isSet('PowerPointApi', '1.4')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'unsupported', message: 'PowerPointApi 1.4 is required to add shapes.' },
      };
    }
    let mutationQueued = false;
    try {
      return await PowerPoint.run(async (ctx) => {
        const op = resolution.op;
        const slide = ctx.presentation.slides.getItem(op.slideId);
        const options: PowerPoint.ShapeAddOptions = {};
        if (op.left !== undefined) options.left = op.left;
        if (op.top !== undefined) options.top = op.top;
        if (op.width !== undefined) options.width = op.width;
        if (op.height !== undefined) options.height = op.height;
        mutationQueued = true;
        const added =
          op.type === 'textBox'
            ? slide.shapes.addTextBox(op.text, options)
            : op.type === 'line'
              ? slide.shapes.addLine(op.connector, options)
              : slide.shapes.addGeometricShape(op.geometry, options);
        if (op.fill !== undefined) added.fill.foregroundColor = op.fill;
        added.load('id');
        await ctx.sync();
        const mintedId = added.id;
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${op.slideId}:${mintedId}`,
          inverse: { op: 'delete-object', objectType: 'shape', name: mintedId },
        };
      });
    } catch {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          'PowerPoint did not confirm the dispatched change. Inspect the slide before trying again.',
        );
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint slide could not be read before adding the shape.',
        },
      };
    }
  }

  private async applyFormatShape(req: ActuationRequest): Promise<ActuationResult> {
    const slideId = req.params.target?.slideId;
    const shapeId = req.params.target?.shapeId;
    const format = req.params.shapeFormat;
    if (!slideId || !shapeId) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'no_target',
          message: 'format-shape needs target.slideId and target.shapeId',
        },
      };
    }
    if (!format) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_format', message: 'format-shape needs params.shapeFormat' },
      };
    }
    if (!isSet('PowerPointApi', '1.4')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'unsupported',
          message: 'PowerPointApi 1.4 is required for shape formatting.',
        },
      };
    }
    if (format.zOrder !== undefined && !isSet('PowerPointApi', '1.8')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'unsupported', message: 'PowerPointApi 1.8 is required for shape z-order.' },
      };
    }

    let mutationQueued = false;
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
        // Capture each prior value just before overwriting it so the recorded inverse holds only
        // what THIS change mutated (restore-shape-format prior keys mirror the params fields).
        const prior: Record<string, string> = {};
        if (format.fill !== undefined) {
          const fill = shape.fill;
          fill.load('foregroundColor');
          await ctx.sync();
          prior['fill'] = String(fill.foregroundColor ?? '');
          mutationQueued = true;
          fill.foregroundColor = format.fill;
        }
        if (format.line !== undefined) {
          const line = shape.lineFormat;
          line.load('color');
          await ctx.sync();
          prior['line'] = String(line.color ?? '');
          mutationQueued = true;
          line.color = format.line;
        }
        if (format.font !== undefined) {
          const font = shape.textFrame.textRange.font;
          font.load('bold,italic,color,size,name,underline');
          await ctx.sync();
          const f = format.font;
          if (f.bold !== undefined) {
            prior['font.bold'] = String(font.bold ?? '');
            mutationQueued = true;
            font.bold = f.bold;
          }
          if (f.italic !== undefined) {
            prior['font.italic'] = String(font.italic ?? '');
            mutationQueued = true;
            font.italic = f.italic;
          }
          if (f.underline !== undefined) {
            prior['font.underline'] = String(font.underline ?? '');
            mutationQueued = true;
            font.underline = f.underline ? 'Single' : 'None';
          }
          if (f.color !== undefined) {
            prior['font.color'] = String(font.color ?? '');
            mutationQueued = true;
            font.color = f.color;
          }
          if (f.size !== undefined) {
            prior['font.size'] = String(font.size ?? '');
            mutationQueued = true;
            font.size = f.size;
          }
          if (f.name !== undefined) {
            prior['font.name'] = String(font.name ?? '');
            mutationQueued = true;
            font.name = f.name;
          }
        }
        if (format.zOrder !== undefined) {
          mutationQueued = true;
          shape.setZOrder(Z_ORDER[format.zOrder]);
        }
        await ctx.sync();
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${slideId}:${shapeId}`,
          inverse: { op: 'restore-shape-format', shapeId, prior },
        };
      });
    } catch {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          'PowerPoint did not confirm the dispatched change. Inspect the slide before trying again.',
        );
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint shape could not be formatted.',
        },
      };
    }
  }

  private async applyAddTableSlide(req: ActuationRequest): Promise<ActuationResult> {
    const slideId = req.params.target?.slideId;
    const grid = req.params.tableGrid;
    if (!slideId) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_target', message: 'add-table-slide needs target.slideId' },
      };
    }
    const columnCount = grid?.rows.reduce((max, row) => Math.max(max, row.length), 0) ?? 0;
    if (!grid || grid.rows.length === 0 || columnCount === 0) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_table', message: 'add-table-slide needs a non-empty tableGrid.rows' },
      };
    }
    if (!isSet('PowerPointApi', '1.8')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'unsupported', message: 'PowerPointApi 1.8 is required for slide tables.' },
      };
    }

    let mutationQueued = false;
    try {
      return await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.slides.getItem(slideId);
        const options: PowerPoint.TableAddOptions = {};
        if (grid.left !== undefined) options.left = grid.left;
        if (grid.top !== undefined) options.top = grid.top;
        if (grid.width !== undefined) options.width = grid.width;
        if (grid.height !== undefined) options.height = grid.height;
        mutationQueued = true;
        const added = slide.shapes.addTable(grid.rows.length, columnCount, options);
        const table = added.getTable();
        grid.rows.forEach((row, r) => {
          row.forEach((value, c) => {
            table.getCellOrNullObject(r, c).text = value;
          });
        });
        added.load('id');
        await ctx.sync();
        const mintedId = added.id;
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${slideId}:${mintedId}`,
          inverse: { op: 'delete-object', objectType: 'shape', name: mintedId },
        };
      });
    } catch {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          'PowerPoint did not confirm the dispatched change. Inspect the slide before trying again.',
        );
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint slide could not be read before adding the table.',
        },
      };
    }
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

/**
 * Geometric shapes {@link PowerPointBridge.actuate} `add-shape` accepts — typed against the
 * literal-union overload of `ShapeCollection.addGeometricShape` (PowerPointApi 1.4, l.184157) so
 * the whitelist doubles as the host parameter type (no casts, nothing outside this set is sent).
 */
const GEOMETRIC_SHAPES = [
  'Rectangle',
  'RoundRectangle',
  'Ellipse',
  'Triangle',
  'Diamond',
  'Pentagon',
  'Hexagon',
  'Octagon',
  'Star5',
  'Chevron',
  'RightArrow',
  'LeftRightArrow',
  'Cloud',
] as const;
type GeometricShapeName = (typeof GEOMETRIC_SHAPES)[number];

/** Contract connector names → the host's `ConnectorType` literals (`addLine`, PowerPointApi 1.4). */
const CONNECTOR_TYPES: Record<'straight' | 'elbow' | 'curve', 'Straight' | 'Elbow' | 'Curve'> = {
  straight: 'Straight',
  elbow: 'Elbow',
  curve: 'Curve',
};

/** Contract z-order names → the host's `ShapeZOrder` literals (`setZOrder`, PowerPointApi 1.8). */
const Z_ORDER: Record<
  'front' | 'back' | 'forward' | 'backward',
  'BringToFront' | 'SendToBack' | 'BringForward' | 'SendBackward'
> = {
  front: 'BringToFront',
  back: 'SendToBack',
  forward: 'BringForward',
  backward: 'SendBackward',
};

interface AddShapeGeometry {
  slideId: string;
  fill?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

/**
 * The fully-resolved host plan for an `add-shape` actuation. The discriminated `type` carries the
 * whitelisted geometry so every branch of the bridge's add call is exhaustively narrowed.
 */
type AddShapeOp =
  | (AddShapeGeometry & { type: 'textBox'; text: string })
  | (AddShapeGeometry & { type: 'line'; connector?: 'Straight' | 'Elbow' | 'Curve' })
  | (AddShapeGeometry & { type: 'geometric'; geometry: GeometricShapeName });

type AddShapeResolution =
  | { ok: true; op: AddShapeOp }
  | {
      ok: false;
      code: 'no_target' | 'no_shape' | 'unsupported';
      message: string;
    };

/**
 * Pure validation for `add-shape`: resolve `params.shape` + `params.target.slideId` into a typed
 * host op, or a precise error code — before any host object is touched.
 */
function resolveAddShape(req: ActuationRequest): AddShapeResolution {
  const slideId = req.params.target?.slideId;
  if (!slideId) {
    return { ok: false, code: 'no_target', message: 'add-shape needs target.slideId' };
  }
  const shape = req.params.shape;
  if (!shape) {
    return { ok: false, code: 'no_shape', message: 'add-shape needs params.shape' };
  }
  const base: AddShapeGeometry = {
    slideId,
    fill: shape.fill,
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
  if (shape.shapeType === 'textBox') {
    return { ok: true, op: { ...base, type: 'textBox', text: shape.text ?? '' } };
  }
  if (shape.shapeType === 'line') {
    return {
      ok: true,
      op: {
        ...base,
        type: 'line',
        connector: shape.connectorType ? CONNECTOR_TYPES[shape.connectorType] : undefined,
      },
    };
  }
  const geometry = GEOMETRIC_SHAPES.find((candidate) => candidate === shape.geometryType);
  if (!geometry) {
    return {
      ok: false,
      code: 'unsupported',
      message: `add-shape geometry "${shape.geometryType ?? ''}" is not supported.`,
    };
  }
  return { ok: true, op: { ...base, type: 'geometric', geometry } };
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

/** Actual dispatch keys; conformance checks these against the advertised capabilities. */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = PowerPointBridge.handledActuations;
