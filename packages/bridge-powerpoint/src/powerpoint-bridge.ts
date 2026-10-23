import { powerpointGeometry, unknownActuationResult } from '@ge/contracts';
import { createBridgeDispatch } from '@ge/runtime';
import type {
  ActuationErrorReason,
  ActuationKind,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  PowerPointGeometry,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { buildDocStateSnapshot } from '@ge/content';
import { POWERPOINT_CAPABILITIES } from './capabilities.js';
import { isSet } from './capabilities-runtime.js';
import {
  mayNameSlide,
  searchSlides,
  selectedShapeToContext,
  selectedSlideToContext,
  shapeContextRef,
  slideContextRef,
  shapesToSlideText,
  slideElementsToDocStateBlocks,
  slideIndexForSelector,
  slideShapeListing,
  slidesToContext,
  tableValuesToText,
  type ShapeElement,
  type SlideElement,
} from './capture.js';
import { planInsertSlide } from './actuate-plan.js';
import { documentChanged, selectionChanged } from './events.js';
import { readPowerPointPage } from './document-reader.js';

/**
 * Upper bound on slides materialized by a single read port (`searchDocument`/
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
  readonly readDocumentPage = readPowerPointPage;

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
      selected.load('items/id');
      await ctx.sync();

      const refs: ContextRef[] = [];
      const first = selected.items[0];
      if (first) {
        const slide = await readSlide(ctx, first);
        refs.push(slideContextRef(slide));
        // Only shapes with text: an empty line or picture has nothing to attach (resolving it gives
        // no context), and listing it gave the model a "shape" to aim a text write at.
        for (const shape of slide.shapes ?? []) {
          if (shape.text.trim()) refs.push(shapeContextRef(slide, shape));
        }
      }
      refs.push({ id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' });
      return refs;
    });
  }

  async resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    if (!isSet('PowerPointApi', '1.4')) return [];
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
      if (!isSet('PowerPointApi', '1.5')) return [];
      return PowerPoint.run(async (ctx) => {
        const selected = ctx.presentation.getSelectedSlides();
        selected.load('items/id');
        await ctx.sync();
        const slide = selected.items[0];
        if (!slide) return [];
        const element = await readSlide(ctx, slide);
        return selectedSlideToContext(element);
      });
    }
    // Whole deck → each slide's shape text → native blocks → chunks (bounded, batched read).
    return slidesToContext('pp:deck', 'Whole deck', await this.readAllSlides());
  }

  /** Monotonic `<doc_state>` version, bumped on each capture (ADR-0003 Layer B element 1). */
  private docStateVersion = 0;

  /**
   * Compact slide inventory without loading shape text. Content lives behind paged document
   * reads or an addressed read. Inventory truncation is explicit; it does not limit traversal.
   * Version increments per capture and is not a source revision.
   */
  documentKey(): string | undefined {
    try {
      return Office.context?.document?.url || undefined;
    } catch {
      return undefined;
    }
  }

  async captureDocState(): Promise<DocStateSnapshot | undefined> {
    if (!isSet('PowerPointApi', '1.4')) return undefined;
    const slides = await this.readAllSlides();
    if (slides.length === 0) return undefined;
    const selected = isSet('PowerPointApi', '1.5')
      ? await PowerPoint.run(async (ctx) => {
          const selection = ctx.presentation.getSelectedSlides();
          selection.load('items/id');
          await ctx.sync();
          return selection.items.map((slide) => {
            // Slide.index needs 1.8. The captured deck order already gives positions on 1.5.
            const index = slides.findIndex((entry) => entry.slideId === slide.id);
            return `${index < 0 ? 'Selected slide' : `Slide ${index + 1}`} (slide:${slide.id})`;
          });
        }).catch(() => []) // Optional selection failure must not discard the captured deck.
      : [];
    this.docStateVersion += 1;
    const snapshot = buildDocStateSnapshot({
      surface: 'powerpoint',
      version: this.docStateVersion,
      blocks: slideElementsToDocStateBlocks(slides),
      ...(selected.length ? { selection: { kind: 'slide', title: selected.join(', ') } } : {}),
    });
    // Shape ids ride on each slide's entry, so shape commands name a real id on the first try
    // instead of guessing and retrying on target_conflict (live 2026-10-02: 3 wasted turns).
    // Read with the slides above: no extra host round trip.
    const shapesBySlide = new Map(
      slides.map((slide) => [`slide:${slide.slideId ?? slide.index}`, slide.shapes ?? []]),
    );
    // Body lines share their slide's locator: summarise once, on the slide's first entry.
    const summarised = new Set<string>();
    for (const entry of snapshot.inventory) {
      if (summarised.has(entry.id)) continue;
      summarised.add(entry.id);
      const summary = shapeSummary(
        entry.id.slice('slide:'.length),
        shapesBySlide.get(entry.id) ?? [],
      );
      // Keep the slide number the builder set ("slide 15"); the shapes follow it.
      if (summary) entry.summary = entry.summary ? `${entry.summary}; ${summary}` : summary;
    }
    return snapshot;
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
    if (!mayNameSlide(selector) || !isSet('PowerPointApi', '1.4')) return [];
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id');
      await ctx.sync();
      const index = slideIndexForSelector(
        slides.items.map((item) => item.id),
        selector,
      );
      const slide = index === undefined ? undefined : slides.items[index];
      if (!slide) return [];
      const [element] = await readSlides(ctx, [slide], [index!]);
      if (!element) return [];
      // A single-slide read is how the model looks before a shape command: list each shape's id,
      // type and which one `shape=title` resolves to, so it can address the shape it means.
      const shapes = ctx.presentation.slides.getItem(slide.id).shapes;
      shapes.load('items/id,items/type');
      await ctx.sync();
      const title = await findTitleShape(ctx, shapes.items).catch(() => undefined);
      const types = new Map(shapes.items.map((shape) => [shape.id, shape.type as string]));
      const listed: SlideElement = {
        ...element,
        shapes: (element.shapes ?? []).map((shape) => ({
          ...shape,
          ...(types.get(shape.shapeId) ? { type: types.get(shape.shapeId) } : {}),
          ...(title?.id === shape.shapeId ? { isTitle: true } : {}),
        })),
      };
      return selectedSlideToContext({
        ...element,
        body: [...element.body, ...slideShapeListing(listed)],
      });
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
   * shared host read behind the whole-deck context, `captureDocState` and `searchDocument`. Bounded
   * so a huge deck can't blow the per-turn budget; read-only (loads shape text, writes nothing).
   */
  private async readAllSlides(): Promise<SlideElement[]> {
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id');
      await ctx.sync();
      const bounded = slides.items.slice(0, MAX_READ_SLIDES);
      return readSlides(
        ctx,
        bounded,
        bounded.map((_, index) => index),
      );
    });
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    return PowerPointBridge.dispatcher.dispatch(this, req);
  }

  private async applySetShapeText(req: ActuationRequest): Promise<ActuationResult> {
    const slideId = req.params.target?.slideId;
    const shapeId = req.params.target?.shapeId;
    const text = paragraphs(req.params.text);
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
        const target = await resolveShapeTarget(ctx, slideId, shapeId);
        if ('error' in target) return slideNotFound(req, target.error, target.reason);
        const slide = ctx.presentation.slides.getItem(target.slideId);
        const shape = slide.shapes.getItemOrNullObject(target.shapeId);
        shape.load('isNullObject,type');
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
              reason: 'shape_gone',
            },
          };
        }
        if (!TEXT_SHAPE_TYPES.has(shape.type)) {
          return slideNotFound(
            req,
            `Shape id ${target.shapeId} is a ${shape.type} and has no text. For the slide title use ` +
              `shape pp:shape:${slideId}:title "…", or the id of a text box.`,
            'no_text',
          );
        }
        const range = shape.textFrame.textRange;
        range.load('text');
        await ctx.sync();
        const priorText = range.text ?? '';
        mutationQueued = true;
        writeParagraphs(range, text);
        await ctx.sync();
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${target.slideId}:${target.shapeId}`,
          inverse: {
            op: 'restore-text',
            anchor: `pp:shape:${target.slideId}:${target.shapeId}`,
            priorText,
          },
        };
      });
    } catch (error) {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the dispatched change${hostErrorSuffix(error)}. ` +
            'Inspect the slide before trying again.',
        );
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint shape could not be re-read before writing.',
          reason: 'reread_failed',
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
    if (plan.base64) {
      const base64 = plan.base64;
      return PowerPoint.run(async (ctx) => {
        // Prebuilt deck path: the agent supplied a Base64 PPTX — let the host merge it (1.2).
        const options = await deckInsertOptions(ctx, plan);
        ctx.presentation.insertSlidesFromBase64(base64, options);
        await ctx.sync();
        const location =
          plan.slideCount === undefined ? 'inserted-deck' : `inserted-deck:${plan.slideCount}`;
        return { ok: true, changeId: req.changeId, kind: req.kind, location };
      });
    }

    // Native compose path: append a slide, then write the title/bullets into it.
    //
    // PowerPoint for the web gives a just-added slide a PROVISIONAL id: writes addressed through
    // that id — `getItem(id)` (5010 "InvalidParam passed to GetItem(id)") or the slide object from
    // the `items` list (GeneralException) — are rejected, while writes by position (`getItemAt`)
    // work; the slide gets its real id once the request completes. So: find the new slide's
    // POSITION by the one id that is new (never a stale index — a co-author's concurrent add must
    // not redirect the text), write to it by position, and read its real id back afterwards.
    // `at=N`: the host only appends, so the slide is moved into place after its text lands. Check
    // the move API before anything is written, so an old host fails cleanly instead of half-done.
    if (plan.targetIndex !== undefined && !isSet('PowerPointApi', '1.8')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'unsupported',
          message:
            'This PowerPoint version cannot place a new slide at a position (needs PowerPointApi 1.8). ' +
            'Add it without at= (it goes at the end) and tell the user to move it.',
        },
      };
    }
    let newIndex: number | undefined;
    let knownIds: NewSlideIds | undefined;
    const result = await PowerPoint.run(async (ctx): Promise<ActuationResult> => {
      try {
        const added = await appendSlide(
          ctx,
          plan.bullets.length > 0 ? CONTENT_LAYOUTS : TITLE_ONLY_LAYOUTS,
        );
        if ('error' in added) return appendFailure(req, added);
        const { index, slideId, shapes, ids } = added;
        await writeComposedSlide(ctx, index, shapes, plan.title, plan.bullets);
        const wrote = queueNewSlideCheck(ctx, index, ids);
        await ctx.sync();
        if (!wrote()) return unknownActuationResult(req, SHIFTED_WHILE_WRITING);
        newIndex = index;
        knownIds = ids;
        // A position past the end means "append": the slide is already there.
        if (plan.targetIndex !== undefined && plan.targetIndex < index) {
          try {
            ctx.presentation.slides.getItemAt(index).moveTo(plan.targetIndex);
            const moved = queueNewSlideCheck(ctx, plan.targetIndex, ids);
            await ctx.sync();
            if (!moved()) return unknownActuationResult(req, SHIFTED_WHILE_WRITING);
            newIndex = plan.targetIndex;
          } catch (error) {
            // The slide and its text landed at the end; only the move is unconfirmed.
            return {
              ok: true,
              changeId: req.changeId,
              kind: req.kind,
              location: `slide:${slideId}`,
              error: {
                code: 'not_moved',
                message:
                  `The slide was added at the end (slide ${index + 1}) but PowerPoint did not confirm ` +
                  `moving it to slide ${plan.targetIndex + 1}${hostErrorSuffix(error)}. ` +
                  'Inspect the deck before trying again.',
              },
            };
          }
        }
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `slide:${slideId}`,
        };
      } catch (error) {
        // The slide may already be in the deck; never report an unconfirmed write as success.
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the new slide and its text${hostErrorSuffix(error)}. ` +
            'Inspect the deck before trying again.',
        );
      }
    });
    if (!result.ok || newIndex === undefined || knownIds === undefined) return result;
    // Best effort: report the slide's settled id (the one `outline` and later commands will see).
    const settledId = await readNewSlideIdAt(newIndex, knownIds);
    return settledId ? { ...result, location: `slide:${settledId}` } : result;
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
        const target = await resolveSlideRef(ctx, op.slideId);
        if ('error' in target) return slideNotFound(req, target.error, target.reason);
        const size = await knownSlideSize(ctx);
        const unplaced = missingGeometryMessage(size, op);
        if (unplaced) return slideNotFound(req, unplaced);
        const offSlide = offSlideMessage(size, op);
        if (offSlide) return slideNotFound(req, offSlide);
        if (op.type === 'textBox') {
          const overPlaceholders = await emptyPlaceholderMessage(ctx, target.slideId, op.slideId);
          if (overPlaceholders) {
            // Nothing written; a later `shape` write into the placeholders settles the task.
            return {
              ok: false,
              changeId: req.changeId,
              kind: req.kind,
              degraded: true,
              error: { code: 'use_placeholders', message: overPlaceholders },
            };
          }
        }
        const slide = ctx.presentation.slides.getItem(target.slideId);
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
        // A geometric shape holds text too ("a blue box saying AI CoE"); it was dropped before.
        if (op.type === 'geometric' && op.text) added.textFrame.textRange.text = op.text;
        added.load('id');
        await ctx.sync();
        const mintedId = added.id;
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${target.slideId}:${mintedId}`,
          inverse: { op: 'delete-object', objectType: 'shape', name: mintedId },
        };
      });
    } catch (error) {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the dispatched change${hostErrorSuffix(error)}. ` +
            'Inspect the slide before trying again.',
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
    if (!format || Object.keys(format).length === 0) {
      // An empty format would "apply" nothing and still report success.
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'no_format',
          message:
            'format-shape needs at least one of fill=, line=, fontColor=, fontSize=, fontBold=, ' +
            "fontItalic=, fontName= or zOrder=. To change a shape's text use: shape " +
            'pp:shape:<slide>:<shape> "text".',
        },
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
        const target = await resolveShapeTarget(ctx, slideId, shapeId);
        if ('error' in target) return slideNotFound(req, target.error, target.reason);
        const slide = ctx.presentation.slides.getItem(target.slideId);
        const shape = slide.shapes.getItemOrNullObject(target.shapeId);
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
              reason: 'shape_gone',
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
          location: `shape:${target.slideId}:${target.shapeId}`,
          inverse: { op: 'restore-shape-format', shapeId: target.shapeId, prior },
        };
      });
    } catch (error) {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the dispatched change${hostErrorSuffix(error)}. ` +
            'Inspect the slide before trying again.',
        );
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint shape could not be formatted.',
          reason: 'format_failed',
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
        error: {
          code: 'no_target',
          message:
            'add-table-slide needs a slide: slide=new title="…" creates a slide for the table; ' +
            `or add it to an existing slide. ${SLIDE_TARGET_HINT}`,
        },
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
    let newSlideIndex: number | undefined;
    let knownIds: NewSlideIds | undefined;
    try {
      const result = await PowerPoint.run(async (ctx): Promise<ActuationResult> => {
        const offSlide = offSlideMessage(await knownSlideSize(ctx), grid);
        if (offSlide) return slideNotFound(req, offSlide);
        let shapes: PowerPoint.ShapeCollection;
        let slideLocation: string;
        /** Where a table goes when the command gave no position: below a title this command wrote. */
        let belowTitle: { left: number; top: number; width: number } | undefined;
        /** On an existing slide: the place worked out from what is already on it. */
        let fitted: TableLayout | undefined;
        if (/^new$/i.test(slideId.trim())) {
          // `slide=new`: this command's name promises a slide, so create one for the table. Write
          // through a fresh position proxy (see appendSlide — PowerPoint web rejects writes addressed
          // by a just-added slide's provisional id), and report the settled id afterwards.
          mutationQueued = true;
          // Title Only: the table goes below the title, not into a content placeholder.
          const added = await appendSlide(ctx, TITLE_ONLY_LAYOUTS);
          if ('error' in added) return appendFailure(req, added);
          newSlideIndex = added.index;
          knownIds = added.ids;
          slideLocation = added.slideId;
          const title = req.params.slide?.title?.trim();
          // The title goes into the layout's title placeholder when there is one, else a text box
          // (same rules as insert-slide).
          if (title) {
            await writeComposedSlide(ctx, added.index, added.shapes, title, []);
            // Without this the host drops the table near the top, over the title (seen live).
            const size = await slideSize(ctx);
            const margin = Math.round(size.width * 0.067);
            belowTitle = {
              left: margin,
              top: Math.round(size.height * 0.27),
              width: size.width - margin * 2,
            };
          }
          shapes = ctx.presentation.slides.getItemAt(added.index).shapes;
        } else {
          if (req.params.slide?.title?.trim()) {
            return slideNotFound(
              req,
              'title= is only used with slide=new (it names the slide that command creates). To ' +
                'retitle an existing slide use: shape pp:shape:<slide>:title "New title".',
            );
          }
          if (typeof req.params.target?.slideIndex === 'number') {
            return slideNotFound(
              req,
              'at= is only used with slide=new (it places the slide that command creates). An ' +
                'existing slide keeps its position.',
            );
          }
          const target = await resolveSlideRef(ctx, slideId);
          if ('error' in target) return slideNotFound(req, target.error, target.reason);
          slideLocation = target.slideId;
          shapes = ctx.presentation.slides.getItem(target.slideId).shapes;
          // A slide that already has content: the table goes below it, never on top of it, and
          // fits on the slide (real PowerPoint, 2026-10-17: a table copied the example's position
          // and covered the bullets, its last row off the slide). A position the command gave
          // is kept only when it is on the slide and clear of every shape.
          fitted = await layoutTableOnSlide(ctx, shapes, grid.rows, columnCount, grid);
          const shrink = fitted?.shrinkBody;
          if (shrink) {
            mutationQueued = true;
            shrink.shape.height = shrink.height;
            await ctx.sync();
            try {
              // The bullets keep their size where they fit; PowerPoint shrinks them where not.
              shrink.shape.textFrame.autoSizeSetting = 'AutoSizeTextToFitShape';
              await ctx.sync();
            } catch {
              // Older hosts: the box is smaller and its text may run under the table.
            }
          }
        }
        const options: PowerPoint.TableAddOptions = fitted
          ? {
              left: fitted.left,
              top: fitted.top,
              width: fitted.width,
              columns: fitted.columns.map((columnWidth) => ({ columnWidth })),
              uniformCellProperties: { font: { size: TABLE_FONT_SIZE } },
            }
          : {};
        if (!fitted) {
          if (grid.left !== undefined) options.left = grid.left;
          else if (belowTitle) options.left = belowTitle.left;
          if (grid.top !== undefined) options.top = grid.top;
          else if (belowTitle) options.top = belowTitle.top;
          if (grid.width !== undefined) options.width = grid.width;
          else if (belowTitle) options.width = belowTitle.width;
          if (grid.height !== undefined) options.height = grid.height;
        }
        // The deck's colour on the header row, not the theme's default (UAT 2026-10-16: purple).
        if (grid.hasHeaders && grid.headerFill) {
          const header = {
            fill: { color: grid.headerFill },
            font: { color: '#FFFFFF', bold: true, ...(fitted ? { size: TABLE_FONT_SIZE } : {}) },
          };
          options.specificCellProperties = grid.rows.map((_, r) =>
            Array.from({ length: columnCount }, () => (r === 0 ? header : {})),
          );
        }
        mutationQueued = true;
        const added = shapes.addTable(grid.rows.length, columnCount, options);
        const table = added.getTable();
        grid.rows.forEach((row, r) => {
          row.forEach((value, c) => {
            table.getCellOrNullObject(r, c).text = lineBreaks(value);
          });
        });
        added.load('id');
        const wrote =
          newSlideIndex !== undefined && knownIds
            ? queueNewSlideCheck(ctx, newSlideIndex, knownIds)
            : undefined;
        await ctx.sync();
        if (wrote && !wrote()) return unknownActuationResult(req, SHIFTED_WHILE_WRITING);
        const mintedId = added.id;
        const landed: ActuationResult = {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${slideLocation}:${mintedId}`,
          inverse: { op: 'delete-object', objectType: 'shape', name: mintedId },
        };
        // `slide=new at=N`: the host only appends, so move the finished slide into place.
        const targetIndex = req.params.target?.slideIndex;
        if (
          newSlideIndex !== undefined &&
          targetIndex !== undefined &&
          targetIndex < newSlideIndex
        ) {
          try {
            ctx.presentation.slides.getItemAt(newSlideIndex).moveTo(targetIndex);
            const moved = knownIds ? queueNewSlideCheck(ctx, targetIndex, knownIds) : undefined;
            await ctx.sync();
            if (moved && !moved()) return unknownActuationResult(req, SHIFTED_WHILE_WRITING);
            newSlideIndex = targetIndex;
          } catch (error) {
            return {
              ...landed,
              error: {
                code: 'not_moved',
                message:
                  `The table slide was added at the end (slide ${newSlideIndex + 1}) but PowerPoint ` +
                  `did not confirm moving it to slide ${targetIndex + 1}${hostErrorSuffix(error)}. ` +
                  'Inspect the deck before trying again.',
              },
            };
          }
        }
        return landed;
      });
      if (!result.ok || newSlideIndex === undefined || knownIds === undefined) return result;
      // The undo deletes this slide, so only the verified new slide's id is accepted here.
      const settledId = await readNewSlideIdAt(newSlideIndex, knownIds);
      if (!settledId) return result;
      // The command created the slide: undoing it means deleting the slide (and the table with it).
      return {
        ...result,
        location: result.location?.replace(/^shape:[^:]+(?=:)/, `shape:${settledId}`),
        inverse: { op: 'delete-object', objectType: 'slide', name: settledId },
      };
    } catch (error) {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the dispatched change${hostErrorSuffix(error)}. ` +
            'Inspect the slide before trying again.',
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
  | (AddShapeGeometry & { type: 'geometric'; geometry: PowerPointGeometry; text?: string });

type AddShapeResolution =
  | { ok: true; op: AddShapeOp }
  | {
      ok: false;
      code: 'no_target' | 'no_shape' | 'no_text' | 'unsupported';
      message: string;
    };

/**
 * Pure validation for `add-shape`: resolve `params.shape` + `params.target.slideId` into a typed
 * host op, or a precise error code — before any host object is touched.
 */
function resolveAddShape(req: ActuationRequest): AddShapeResolution {
  const slideId = req.params.target?.slideId;
  if (!slideId) {
    return {
      ok: false,
      code: 'no_target',
      message: `add-shape needs target.slideId. ${SLIDE_TARGET_HINT}`,
    };
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
    const text = shape.text?.trim() ?? '';
    if (!text) {
      // Live: a malformed `/add-shape slide:2 {"text": …}` arrived with no text and was "applied"
      // as an empty box in the corner.
      return {
        ok: false,
        code: 'no_text',
        message:
          'add-shape shapeType=textBox needs text="…". Example: /add-shape slide=2 shapeType=textBox ' +
          'text="Draft" left=72 top=330 width=400 height=40',
      };
    }
    return { ok: true, op: { ...base, type: 'textBox', text: paragraphs(shape.text) ?? '' } };
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
  // The contracts' whitelist doubles as the host parameter type: nothing outside it is sent.
  const geometry = powerpointGeometry(shape.geometryType);
  if (!geometry) {
    return {
      ok: false,
      code: 'unsupported',
      message: `add-shape geometry "${shape.geometryType ?? ''}" is not supported.`,
    };
  }
  const text = paragraphs(shape.text?.trim());
  return { ok: true, op: { ...base, type: 'geometric', geometry, ...(text ? { text } : {}) } };
}

/**
 * The model writes line breaks as a literal `\n` inside a command argument; the slide showed the
 * two characters (UAT 2026-10-16). Make them real breaks.
 */
function lineBreaks<T extends string | undefined>(text: T): T {
  return (text?.replace(/(?:\\r)?\\n/g, '\n') ?? text) as T;
}

/**
 * Text for a whole shape (shape text, text boxes, shapes with text): every line — a literal `\n`
 * the model wrote or a real newline — becomes its own paragraph (`\r`). A `\n` is only a line
 * break inside ONE paragraph, so a typed list rendered as a single run-on paragraph with no list
 * formatting (UAT 2026-10-16, "daftar isi" on a slide). Line breaks inside a title or a single
 * bullet stay {@link lineBreaks}.
 */
function paragraphs<T extends string | undefined>(text: T): T {
  return (lineBreaks(text)?.replace(/\r?\n/g, PARAGRAPH_BREAK) ?? text) as T;
}

/**
 * Lines that carry their own list markers ("1. Latar belakang", "a) …", "- …"): at least two, and
 * every non-empty one starts with one. A body placeholder adds its bullet in front of each, giving
 * "• 1. Latar belakang" (UAT 2026-10-16) — so for these the layout's bullet is hidden.
 */
function hasOwnListMarkers(lines: readonly string[]): boolean {
  const filled = lines.filter((line) => line.trim() !== '');
  return filled.length >= 2 && filled.every((line) => OWN_LIST_MARKER.test(line));
}
const OWN_LIST_MARKER = /^\s*(?:(?:\d{1,3}|[a-z]|[ivxlc]{1,6})[.)]|[•\-–*])\s/i;

/**
 * Write `text` into a shape's text range, one paragraph per line; when the lines number
 * themselves, hide the bullet the placeholder would add (`bulletFormat`, PowerPointApi 1.4).
 */
function writeParagraphs(range: PowerPoint.TextRange, text: string): void {
  range.text = text;
  if (hasOwnListMarkers(text.split(PARAGRAPH_BREAK)) && isSet('PowerPointApi', '1.4')) {
    range.paragraphFormat.bulletFormat.visible = false;
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

/**
 * Shape types whose `textFrame` is readable. Per the typings, `Shape.textFrame` throws
 * `InvalidArgument` for a shape without one (image, table, chart, group, media, ...), and a throw
 * fails the whole batched sync — so only these types are asked for text.
 */
const TEXT_SHAPE_TYPES: ReadonlySet<string> = new Set([
  'GeometricShape',
  'TextBox',
  'Placeholder',
  'Callout',
  'Freeform',
]);

/**
 * Shapes whose content is read: text-frame shapes, plus native tables (via `Table.values`,
 * PowerPointApi 1.8) where the host supports them. Pictures and charts have no text API.
 */
function isReadableShape(shape: PowerPoint.Shape): boolean {
  if (shape.type === 'Table') return isSet('PowerPointApi', '1.8');
  return TEXT_SHAPE_TYPES.has(shape.type);
}

/**
 * A placeholder reports `type: 'Placeholder'` whatever it holds; one filled with a picture, table
 * or chart has no text frame. `placeholderFormat.containedType` (PowerPointApi 1.8) is `null` for an
 * empty or text placeholder, else the contained shape's type.
 */
function placeholderHoldsText(containedType: string | null | undefined): boolean {
  return (
    containedType === null ||
    containedType === undefined ||
    (containedType !== 'Placeholder' && TEXT_SHAPE_TYPES.has(containedType))
  );
}

/** Read one slide's shapes' text (+ id/index) into a pure {@link SlideElement}. */
async function readSlide(
  ctx: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
): Promise<SlideElement> {
  const deck = ctx.presentation.slides;
  deck.load('items/id');
  slide.load('id');
  await ctx.sync();
  const index = deck.items.findIndex((item) => item.id === slide.id);
  if (index < 0) throw new Error('PowerPoint slide no longer exists.');
  const [element] = await readSlides(ctx, [slide], [index]);
  if (!element) throw new Error('PowerPoint slide read returned no slide.');
  return element;
}

/**
 * Read slides' shape text in a constant number of host round-trips (at most three syncs, whatever
 * the slide count). Each sync is a full round-trip on PowerPoint for the web (~0.85s measured), so a
 * per-slide sync loop made a 9-slide whole-deck attach take ~16s; batched, the same read takes ~1.5s.
 */
async function readSlides(
  ctx: PowerPoint.RequestContext,
  slides: PowerPoint.Slide[],
  positions: number[],
): Promise<SlideElement[]> {
  if (slides.length === 0) return [];
  const shapeCollections = slides.map((slide) => {
    slide.load('id');
    const shapes = slide.shapes;
    shapes.load('items/id,items/type');
    return shapes;
  });
  await ctx.sync();

  const candidates = shapeCollections.flatMap((shapes) => shapes.items.filter(isReadableShape));
  const readable = await withoutNonTextPlaceholders(ctx, candidates);
  const texts = await readShapeContent(ctx, readable);

  return slides.map((slide, slideIndex) => {
    const shapes = shapeCollections[slideIndex]?.items ?? [];
    const shapeTexts = shapes.map((shape) => texts.get(shape) ?? '');
    const { title, body } = shapesToSlideText(shapeTexts);
    return {
      index: positions[slideIndex]!,
      slideId: slide.id,
      title,
      body,
      shapes: shapes.map((shape, index) => ({
        shapeId: shape.id,
        text: shapeTexts[index] ?? '',
      })),
    };
  });
}

/**
 * Drop placeholders holding a picture/table/chart (one sync, only when the batch has placeholders).
 * Without PowerPointApi 1.8, or if the host refuses the lookup, placeholders are kept and
 * {@link readShapeContent}'s per-shape fallback skips any that throw.
 */
async function withoutNonTextPlaceholders(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<PowerPoint.Shape[]> {
  const placeholders = shapes.filter((shape) => shape.type === 'Placeholder');
  if (placeholders.length === 0 || !isSet('PowerPointApi', '1.8')) return shapes;
  try {
    const formats = new Map(
      placeholders.map((shape) => {
        const format = shape.placeholderFormat;
        format.load('containedType');
        return [shape, format] as const;
      }),
    );
    await ctx.sync();
    return shapes.filter((shape) => {
      const format = formats.get(shape);
      return !format || placeholderHoldsText(format.containedType);
    });
  } catch {
    return shapes;
  }
}

/**
 * Queue a read of one shape's content: a table's cell grid, else its text frame. Returns a getter to
 * call after the next `ctx.sync()`.
 */
function queueShapeRead(shape: PowerPoint.Shape): () => string {
  if (shape.type === 'Table') {
    const table = shape.getTable();
    table.load('values');
    return () => tableValuesToText(table.values ?? []);
  }
  const range = shape.textFrame.textRange;
  range.load('text');
  return () => range.text ?? '';
}

/**
 * Read `shapes`' content (text frames and tables) in one sync. A single shape the host refuses
 * fails the whole batch, so on failure fall back to one sync per shape and skip the refusals —
 * slower, but one unexpected shape can't block the read of the rest of the deck.
 */
async function readShapeContent(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<Map<PowerPoint.Shape, string>> {
  const texts = new Map<PowerPoint.Shape, string>();
  if (shapes.length === 0) return texts;
  try {
    const reads = shapes.map((shape) => [shape, queueShapeRead(shape)] as const);
    await ctx.sync();
    for (const [shape, read] of reads) texts.set(shape, read());
    return texts;
  } catch {
    for (const shape of shapes) {
      try {
        const read = queueShapeRead(shape);
        await ctx.sync();
        texts.set(shape, read());
      } catch {
        // Nothing readable on this shape: leave it empty.
      }
    }
    return texts;
  }
}

async function readShapeContext(
  ctx: PowerPoint.RequestContext,
  target: PowerPointRevealTarget,
): Promise<ResolvedContext[]> {
  const slide = ctx.presentation.slides.getItem(target.slideId);
  const deck = ctx.presentation.slides;
  deck.load('items/id');
  slide.load('id');
  const shape = slide.shapes.getItemOrNullObject(target.shapeId ?? '');
  shape.load('id,isNullObject,type');
  await ctx.sync();
  if (shape.isNullObject || !target.shapeId || !isReadableShape(shape)) return [];
  const index = deck.items.findIndex((item) => item.id === slide.id);
  if (index < 0) return [];
  const [readable] = await withoutNonTextPlaceholders(ctx, [shape]);
  if (!readable) return [];
  const texts = await readShapeContent(ctx, [readable]);
  return selectedShapeToContext(
    { index, slideId: slide.id },
    { shapeId: shape.id, text: texts.get(readable) ?? '' },
  );
}

/** Placeholder types that take a slide title / its body text, in preference order. */
const TITLE_PLACEHOLDERS: readonly string[] = ['Title', 'CenterTitle', 'VerticalTitle'];
// Not 'Subtitle': a cover layout's subtitle is small, light and centred — bullets written there were
// unreadable (UAT 2026-10-16). With no body placeholder they go in a left-aligned text box.
const BODY_PLACEHOLDERS: readonly string[] = ['Body', 'Content', 'VerticalBody'];

/**
 * Separates bullets: `\r` is a paragraph break (what PowerPoint reads back for typed paragraphs),
 * `\n` only a line break inside one paragraph — joined with `\n`, every bullet after the first
 * rendered as a continuation of the first, with no bullet of its own (live 2026-10-16).
 */
const PARAGRAPH_BREAK = '\r';

/**
 * Slide size used for fallback text boxes when the host can't report it (`pageSetup` is
 * PowerPointApi 1.10). 720 × 405 pt is the smallest standard slide (10" × 5.625"), so boxes laid
 * out for it stay on-slide on 16:9 widescreen (960 × 540) and 4:3 (720 × 540) decks too.
 */
const FALLBACK_SLIDE_SIZE = { width: 720, height: 405 };

/**
 * Write a composed title + bullets into a freshly added slide.
 *
 * `slides.add()` uses the deck's default layout, and a layout may have no title/body placeholders at
 * all (a "Blank" layout, or decks generated by tools such as pptxgenjs). So: write into the title and
 * body placeholders when the slide has them (matched by placeholder type on PowerPointApi 1.8, else
 * the first two placeholders in order); for any part with no placeholder, add a sized text box.
 * Every requested part is written somewhere — never a silent no-op.
 *
 * `shapes` (from a loaded probe) is only READ, for types and positions. Every WRITE goes through a
 * fresh, never-loaded `slides.getItemAt(slideIndex)` (and `shapes.getItemAt(i)`) proxy: PowerPoint
 * for the web rejects writes addressed through a just-added slide's provisional id (5010
 * "InvalidParam passed to GetItem(id)"), and a loaded proxy is addressed by id.
 */
async function writeComposedSlide(
  ctx: PowerPoint.RequestContext,
  slideIndex: number,
  shapes: PowerPoint.Shape[],
  rawTitle: string,
  rawBullets: string[],
): Promise<void> {
  const title = lineBreaks(rawTitle);
  const bullets = rawBullets.map(lineBreaks);
  const { titleShape, bodyShape } = await findTextPlaceholders(ctx, shapes);
  const needsBox = (title && !titleShape) || (bullets.length > 0 && !bodyShape);
  const size = needsBox ? await slideSize(ctx) : FALLBACK_SLIDE_SIZE;
  const margin = Math.round(size.width * 0.067);
  const width = size.width - margin * 2;
  const freshShapes = (): PowerPoint.ShapeCollection =>
    ctx.presentation.slides.getItemAt(slideIndex).shapes;
  const rangeOf = (shape: PowerPoint.Shape): PowerPoint.TextRange =>
    freshShapes().getItemAt(shapes.indexOf(shape)).textFrame.textRange;
  const writeInto = (shape: PowerPoint.Shape, text: string): void => {
    rangeOf(shape).text = text;
  };

  if (title) {
    if (titleShape) {
      writeInto(titleShape, title);
    } else {
      const box = freshShapes().addTextBox(title, {
        left: margin,
        top: Math.round(size.height * 0.07),
        width,
        height: Math.round(size.height * 0.18),
      });
      box.textFrame.textRange.font.size = 32;
      box.textFrame.textRange.font.bold = true;
    }
  }
  if (bullets.length > 0) {
    if (bodyShape) {
      // A body placeholder brings the layout's own bullet style — one paragraph per bullet.
      writeParagraphs(rangeOf(bodyShape), bullets.join(PARAGRAPH_BREAK));
    } else {
      const numbered = hasOwnListMarkers(bullets);
      const box = freshShapes().addTextBox(
        bullets.map((bullet) => (numbered ? bullet : `• ${bullet}`)).join(PARAGRAPH_BREAK),
        {
          left: margin,
          top: Math.round(size.height * 0.28),
          width,
          height: Math.round(size.height * 0.62),
        },
      );
      box.textFrame.textRange.font.size = 18;
    }
  }
}

/** The new slide's title and body placeholders, if its layout has them. */
async function findTextPlaceholders(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<{ titleShape?: PowerPoint.Shape; bodyShape?: PowerPoint.Shape }> {
  const placeholders = shapes.filter((shape) => shape.type === 'Placeholder');
  if (placeholders.length === 0) return {};
  if (!isSet('PowerPointApi', '1.8')) {
    // No placeholder types to match on: title first, body second (the layout convention).
    return { titleShape: placeholders[0], bodyShape: placeholders[1] };
  }
  const formats = placeholders.map((shape) => {
    const format = shape.placeholderFormat;
    format.load('type');
    return [shape, format] as const;
  });
  await ctx.sync();
  const firstOf = (types: readonly string[]): PowerPoint.Shape | undefined => {
    for (const type of types) {
      const hit = formats.find(([, format]) => format.type === type);
      if (hit) return hit[0];
    }
    return undefined;
  };
  const titleShape = firstOf(TITLE_PLACEHOLDERS);
  const bodyShape = firstOf(BODY_PLACEHOLDERS);
  return { ...(titleShape ? { titleShape } : {}), ...(bodyShape ? { bodyShape } : {}) };
}

/** Cell text size for a table placed on a slide that already has content. */
const TABLE_FONT_SIZE = 14;

/** Where a table goes on a slide that already has content. */
interface TableLayout {
  left: number;
  top: number;
  width: number;
  columns: number[];
  /** The body placeholder made shorter so the table fits below it. */
  shrinkBody?: { shape: PowerPoint.Shape; height: number };
}

/**
 * Column widths that add up to `width`: a short column ("No", "#") gets just enough room, the
 * others share the rest by their longest text.
 */
export function tableColumnWidths(rows: string[][], columnCount: number, width: number): number[] {
  const longest = Array.from({ length: columnCount }, (_, c) =>
    Math.min(60, Math.max(2, ...rows.map((row) => (row[c] ?? '').length))),
  );
  const narrow = longest.map((n) => n <= 4);
  const narrowWidth = Math.min(56, width / columnCount);
  const flexible = width - narrowWidth * narrow.filter(Boolean).length;
  const flexTotal = longest.reduce((sum, n, c) => sum + (narrow[c] ? 0 : n), 0);
  return longest.map((n, c) =>
    narrow[c] ? narrowWidth : flexTotal > 0 ? (flexible * n) / flexTotal : flexible,
  );
}

/** Rough table height at {@link TABLE_FONT_SIZE}: each row as tall as its most wrapped cell. */
export function tableHeight(rows: string[][], widths: number[]): number {
  const charWidth = TABLE_FONT_SIZE * 0.5;
  const lineHeight = TABLE_FONT_SIZE * 1.2;
  return rows.reduce((height, row) => {
    const lines = Math.max(
      1,
      ...row.map((text, c) =>
        Math.ceil((text.length * charWidth) / Math.max(20, (widths[c] ?? 100) - 14.4)),
      ),
    );
    return height + 7.2 + lineHeight * lines;
  }, 0);
}

/**
 * Lay a table out below what a slide already holds: under the title, under the body (made shorter
 * when the table would not fit beneath it, never below a third of the room), and under any other
 * shape. Full content width, slide margins kept.
 */
async function layoutTableOnSlide(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  rows: string[][],
  columnCount: number,
  asked: { left?: number; top?: number; width?: number; height?: number },
): Promise<TableLayout | undefined> {
  shapes.load('items/id,items/type,items/left,items/top,items/width,items/height');
  await ctx.sync();
  const size = await slideSize(ctx);
  const box = (shape: PowerPoint.Shape) => {
    const n = (v: number | undefined) => (Number.isFinite(v) ? (v as number) : 0);
    return {
      left: n(shape.left),
      top: n(shape.top),
      width: n(shape.width),
      height: n(shape.height),
    };
  };
  const margin = Math.round(size.width * 0.067);
  const width = size.width - margin * 2;
  const columns = tableColumnWidths(rows, columnCount, width);
  const needed = tableHeight(rows, columns);
  // A position the command gave is kept when it is on the slide and clear of every shape.
  if (asked.left !== undefined && asked.top !== undefined) {
    const mine = {
      left: asked.left,
      top: asked.top,
      width: asked.width ?? width,
      height: asked.height ?? needed,
    };
    const clear = shapes.items.every((shape) => {
      const b = box(shape);
      return (
        b.left + b.width <= mine.left ||
        mine.left + mine.width <= b.left ||
        b.top + b.height <= mine.top ||
        mine.top + mine.height <= b.top
      );
    });
    if (clear && mine.left + mine.width <= size.width && mine.top + mine.height <= size.height)
      return undefined;
  }
  const { titleShape, bodyShape } = await findTextPlaceholders(ctx, shapes.items);
  const bottom = size.height - Math.round(margin / 2);
  let top = titleShape ? box(titleShape).top + box(titleShape).height + 12 : margin;
  let shrinkBody: TableLayout['shrinkBody'];
  if (bodyShape) {
    const body = box(bodyShape);
    const bodyTop = Math.max(body.top, top - 12);
    const room = bottom - bodyTop;
    const height = Math.min(body.height, Math.max(room / 3, room - needed - 12));
    if (height < body.height) shrinkBody = { shape: bodyShape, height: Math.round(height) };
    top = bodyTop + height + 12;
  }
  for (const shape of shapes.items) {
    if (shape === titleShape || shape === bodyShape) continue;
    const b = box(shape);
    top = Math.max(top, b.top + b.height + 12);
  }
  return {
    left: margin,
    top: Math.round(top),
    width,
    columns,
    ...(shrinkBody ? { shrinkBody } : {}),
  };
}

/** The deck's slide size in points, or {@link FALLBACK_SLIDE_SIZE} if the host can't report it. */
async function slideSize(
  ctx: PowerPoint.RequestContext,
): Promise<{ width: number; height: number }> {
  if (!isSet('PowerPointApi', '1.10')) return FALLBACK_SLIDE_SIZE;
  try {
    const setup = ctx.presentation.pageSetup;
    setup.load('slideWidth,slideHeight');
    await ctx.sync();
    if (setup.slideWidth > 0 && setup.slideHeight > 0) {
      return { width: setup.slideWidth, height: setup.slideHeight };
    }
  } catch {
    // Fall through to the size that fits every standard slide.
  }
  return FALLBACK_SLIDE_SIZE;
}

/**
 * Resolve the slide a write command names. Commands come from the model, which may give an exact
 * host id (`256#0` on PowerPoint web), `slide:<id>` / `pp:slide:<id>`, a 1-based slide number, or `last` for the
 * deck's last slide (e.g. the one a preceding `insert-slide` just added — the planner makes a
 * `slide=last` write depend on that insert, so it never runs if the insert failed). Resolved with one
 * read-only sync BEFORE any mutation is queued, so a bad reference fails cleanly with the valid
 * choices instead of surfacing as an "outcome uncertain" write.
 */
async function resolveSlideRef(
  ctx: PowerPoint.RequestContext,
  raw: string,
): Promise<{ slideId: string } | { error: string; reason: ActuationErrorReason }> {
  const slides = ctx.presentation.slides;
  slides.load('items/id');
  await ctx.sync();
  const ids = slides.items.map((slide) => slide.id);
  const key = raw.trim().replace(/^(?:pp:)?slide:/i, '');
  if (/^last$/i.test(key)) {
    const last = ids.at(-1);
    return last ? { slideId: last } : { error: 'The deck has no slides yet.', reason: 'no_slides' };
  }
  if (ids.includes(key)) return { slideId: key };
  if (/^\d+$/.test(key)) {
    const id = ids[Number(key) - 1];
    if (id) return { slideId: id };
  }
  const shown = ids.slice(0, 12).join(', ') + (ids.length > 12 ? ', …' : '');
  return {
    error:
      `Slide ${JSON.stringify(raw.slice(0, 64))} was not found. Use a slide id from the outline (${shown}), ` +
      `a slide number 1–${ids.length}, or "last". To create a new slide with a title and bullets, ` +
      'use the slide command: slide "Title" "bullet" "bullet".',
    reason: 'slide_not_found',
  };
}

/**
 * Appended to a missing-slide error so the model can recover in one turn: models reach for the
 * `/add-*` commands to CREATE a slide with text; `slide "Title" "bullet"` is the command for that.
 */
const SLIDE_TARGET_HINT =
  'Pass slide= a slide id from the outline, a slide number, or last. To create a new slide with a title and ' +
  'bullets, use the slide command instead: slide "Title" "bullet" "bullet".';

/**
 * Append a slide and identify it. Finds the new slide's POSITION by the one id that is new (never a
 * pre-computed index — a co-author's concurrent add must not redirect later writes), then reads its
 * shapes through a probe proxy. The caller must WRITE through fresh, never-loaded
 * `slides.getItemAt(index)` proxies: PowerPoint for the web gives a just-added slide a provisional
 * id that rejects writes (5010 "InvalidParam passed to GetItem(id)"), and loading a proxy
 * re-addresses it by id. Queues the add — callers treat any failure after this as outcome-unknown.
 */
async function appendSlide(
  ctx: PowerPoint.RequestContext,
  layoutTypes: readonly string[],
): Promise<
  | {
      index: number;
      slideId: string;
      shapes: PowerPoint.Shape[];
      /** The new slide's identity — what later checks prove "still the new slide" against. */
      ids: NewSlideIds;
    }
  | { error: string; removed?: true }
> {
  const slides = ctx.presentation.slides;
  slides.load('items/id');
  await ctx.sync();
  const existingIds = new Set(slides.items.map((item) => item.id));
  const layoutId = await findLayoutId(ctx, existingIds.size, layoutTypes);
  slides.add(layoutId ? { layoutId } : undefined);
  await ctx.sync();
  const after = ctx.presentation.slides;
  after.load('items/id');
  await ctx.sync();
  const newIndexes = after.items.flatMap((item, index) =>
    existingIds.has(item.id) ? [] : [index],
  );
  const index = newIndexes.length === 1 ? newIndexes[0] : undefined;
  const provisionalId = index === undefined ? undefined : after.items[index]?.id;
  if (index === undefined || provisionalId === undefined) {
    return {
      error:
        'PowerPoint added a slide but it could not be identified. Inspect the deck before trying again.',
    };
  }
  // PowerPoint for the web finishes a new slide asynchronously (live 2026-10-16): for about a second
  // it keeps a provisional id, and a slide added with a layout has NO shapes yet — text written then
  // falls back to text boxes, or is lost. So when a layout was asked for (it has placeholders), wait
  // for its shapes. Desktop hosts have them at once and never wait.
  //
  // Each probe re-reads every id too (queueNewSlideCheck): the new slide is "the one id that wasn't
  // there before", and its id may swap from provisional to real ONCE.
  const ids: NewSlideIds = { before: existingIds, provisional: provisionalId };
  for (let attempt = 0; ; attempt++) {
    const probe = ctx.presentation.slides.getItemAt(index);
    probe.load('id');
    const shapes = probe.shapes;
    shapes.load('items/id,items/type');
    const check = queueNewSlideCheck(ctx, index, ids);
    await ctx.sync();
    const id = check();
    if (id === undefined || id !== probe.id) {
      return {
        error:
          'The deck changed while the slide was being added. Inspect the deck before trying again.',
      };
    }
    if (!layoutId || shapes.items.length > 0) {
      return { index, slideId: id, shapes: shapes.items, ids };
    }
    if (attempt >= NEW_SLIDE_SETTLE_ATTEMPTS) {
      // PowerPoint never finished the slide, so nothing was written to it. Remove the empty slide
      // we just added, so trying again cannot leave a duplicate (security review 2026-10-17).
      if (await removeNewSlide(ctx, index, ids)) {
        return {
          error: 'PowerPoint did not finish creating the slide in time, so nothing was added.',
          removed: true,
        };
      }
      return {
        error:
          'PowerPoint added the slide but did not finish creating it in time, so its text was not ' +
          'written. Inspect the deck before trying again.',
      };
    }
    await new Promise((resolve) => setTimeout(resolve, NEW_SLIDE_SETTLE_POLL_MS));
  }
}

/**
 * Queue a read that confirms the slide at `index` is still the ONE slide that was not in the deck
 * before the add — in the same batch as the writes it guards, so the writes provably hit that
 * slide (a batch runs in order). After the sync, the returned function gives that slide's id, or
 * undefined if a co-author added, removed or moved a slide in between.
 */
function queueNewSlideCheck(
  ctx: PowerPoint.RequestContext,
  index: number,
  ids: NewSlideIds,
): () => string | undefined {
  const slide = ctx.presentation.slides.getItemAt(index);
  slide.load('id');
  const all = ctx.presentation.slides;
  all.load('items/id');
  return () => {
    const now = all.items.map((item) => item.id);
    const fresh = now.filter((id) => !ids.before.has(id));
    if (now.length !== ids.before.size + 1 || fresh.length !== 1 || fresh[0] !== slide.id)
      return undefined;
    // After the one swap only the real id is ours; a later different id (or a swap back) is not.
    if (ids.real !== undefined) return slide.id === ids.real ? slide.id : undefined;
    if (slide.id !== ids.provisional) ids.real = slide.id;
    return slide.id;
  };
}

/**
 * A slide this request added: the ids before the add, its first (provisional) id, and the real id
 * PowerPoint for the web may swap it to once. Every later check accepts only this slide.
 */
interface NewSlideIds {
  readonly before: ReadonlySet<string>;
  readonly provisional: string;
  real?: string;
}

/** Delete the empty slide this request just added; true only when the deck is back as it was. */
async function removeNewSlide(
  ctx: PowerPoint.RequestContext,
  index: number,
  ids: NewSlideIds,
): Promise<boolean> {
  try {
    const check = queueNewSlideCheck(ctx, index, ids);
    await ctx.sync();
    if (check() === undefined) return false;
    ctx.presentation.slides.getItemAt(index).delete();
    const all = ctx.presentation.slides;
    all.load('items/id');
    await ctx.sync();
    const now = all.items.map((item) => item.id);
    return now.length === ids.before.size && now.every((id) => ids.before.has(id));
  } catch {
    return false;
  }
}

/** How long {@link appendSlide} waits for a new layout slide's shapes: 20 × 200 ms = 4 s. */
const NEW_SLIDE_SETTLE_ATTEMPTS = 20;
const NEW_SLIDE_SETTLE_POLL_MS = 200;

/** A failed add: a clean failure when the empty slide was removed again, else outcome unknown. */
function appendFailure(
  req: ActuationRequest,
  added: { error: string; removed?: true },
): ActuationResult {
  if (!added.removed) return unknownActuationResult(req, added.error);
  return {
    ok: false,
    changeId: req.changeId,
    kind: req.kind,
    error: { code: 'slide_not_ready', message: added.error },
  };
}

const SHIFTED_WHILE_WRITING =
  'The deck changed while the new slide was being written, so the text may have landed on another ' +
  'slide. Inspect the deck before trying again.';

/**
 * Layouts a new slide asks for, by `SlideLayout.type` in preference order. With no layout,
 * `slides.add()` takes the master's FIRST layout — "Title Slide" in most decks — and the bullets
 * land in its centred 24 pt subtitle (UAT 2026-10-16). Types, not names: layout names are
 * localised ("Judul dan Konten") and renamed by templates.
 */
const CONTENT_LAYOUTS: readonly string[] = ['Object', 'Text'];
const TITLE_ONLY_LAYOUTS: readonly string[] = ['TitleOnly'];

/** The layout chosen per slide kind and deck layout set, so later slides skip the shape reads. */
const chosenLayouts = new Map<string, string | null>();

/**
 * The id of the layout a new slide should use, or undefined (then the host's default applies, as
 * before). An appended slide takes the previous slide's master, so the layout comes from the last
 * slide's master (the first master for an empty deck) — another master's layout is ignored by
 * PowerPoint for the web (seen live: it silently adds the default layout).
 *
 * A layout fits by its placeholders: a content slide needs a title and a body placeholder; a
 * title-only slide a `Title` placeholder and no body or subtitle (so never the cover). Among the
 * fits, `types` (`SlideLayout.type`, in preference order) wins, then the template's own name
 * (English or Indonesian Office), then the first fit.
 *
 * Read through position proxies only: PowerPoint for the web re-addresses a loaded proxy by id,
 * and master/layout lookup by id fails (5010 "InvalidParam passed to GetItem(id)", live
 * 2026-10-16). Read-only and best effort — a failed read never blocks the slide.
 */
async function findLayoutId(
  ctx: PowerPoint.RequestContext,
  slideCount: number,
  types: readonly string[],
): Promise<string | undefined> {
  if (types.length === 0 || !isSet('PowerPointApi', '1.8')) return undefined;
  const master = (): PowerPoint.SlideMaster =>
    slideCount > 0
      ? ctx.presentation.slides.getItemAt(slideCount - 1).slideMaster
      : ctx.presentation.slideMasters.getItemAt(0);
  const titleOnly = types === TITLE_ONLY_LAYOUTS;
  try {
    const layouts = master().layouts;
    layouts.load('items/id,items/name');
    await ctx.sync();
    // Keyed by the deck's own layout ids: ids alone repeat across presentations.
    const key = [types.join(','), ...layouts.items.map((layout) => layout.id)].join('|');
    const cached = chosenLayouts.get(key);
    if (cached !== undefined) return cached ?? undefined;
    const typed = layouts.items.map((_, i) => {
      const layout = master().layouts.getItemAt(i);
      layout.load('type');
      const shapes = layout.shapes;
      shapes.load('items/type');
      return { layout, shapes };
    });
    await ctx.sync();
    const formats = typed.map(({ shapes }, i) =>
      shapes.items.flatMap((shape, j) => {
        if (shape.type !== 'Placeholder') return [];
        const format = master().layouts.getItemAt(i).shapes.getItemAt(j).placeholderFormat;
        format.load('type');
        return [format];
      }),
    );
    await ctx.sync();
    const fits = typed.flatMap(({ layout }, i) => {
      const kinds = new Set(formats[i]?.map((format) => String(format.type)));
      const body = BODY_PLACEHOLDERS.some((type) => kinds.has(type));
      const fit = titleOnly
        ? kinds.has('Title') && !body && !kinds.has('Subtitle')
        : TITLE_PLACEHOLDERS.some((type) => kinds.has(type)) && body;
      const id = layouts.items[i]?.id;
      return fit && id
        ? [{ id, type: String(layout.type), name: layouts.items[i]?.name ?? '' }]
        : [];
    });
    const named = titleOnly ? /title\s*only|hanya\s*judul/i : /title.*content|judul.*(konten|isi)/i;
    const chosen =
      types.map((type) => fits.find((fit) => fit.type === type)).find(Boolean) ??
      fits.find((fit) => named.test(fit.name)) ??
      fits[0];
    chosenLayouts.set(key, chosen?.id ?? null);
    return chosen?.id;
  } catch {
    // Fall through to the host's default layout.
  }
  return undefined;
}

/**
 * Resolve the slide + shape a shape command names. Slides as in {@link resolveSlideRef}; shapes by
 * exact host id or `title` (the Title/CenterTitle placeholder, else the first shape with text).
 * Deliberately no shape ordinals (see below). Read-only, before any mutation; an unknown reference
 * returns the valid choices.
 */
async function resolveShapeTarget(
  ctx: PowerPoint.RequestContext,
  slideRef: string,
  shapeRef: string,
): Promise<{ slideId: string; shapeId: string } | { error: string; reason: ActuationErrorReason }> {
  const slide = await resolveSlideRef(ctx, slideRef);
  if ('error' in slide) return slide;
  const collection = ctx.presentation.slides.getItem(slide.slideId).shapes;
  collection.load('items/id,items/type');
  await ctx.sync();
  const shapes = collection.items;
  const key = shapeRef.trim().replace(/^(?:pp:)?shape:/i, '');
  const exact = shapes.find((shape) => shape.id === key);
  if (exact) return { slideId: slide.slideId, shapeId: exact.id };
  if (/^title$/i.test(key)) {
    const title = await findTitleShape(ctx, shapes);
    if (title) return { slideId: slide.slideId, shapeId: title.id };
  }
  // No shape ordinals: PowerPoint shape ids are small numbers themselves (2, 3, 4…), so "2" would be
  // ambiguous between "id 2" and "the 2nd shape". Only exact ids and `title` resolve.
  const shown = await describeShapes(ctx, shapes.slice(0, 12));
  // Spell out both ref forms: the `shape` command takes `pp:shape:<slide>:<shape>`, the `/…-shape`
  // kinds take `shape=`. A model told only "shape=title" retries the `shape` command blind.
  const slideKey = slideRef
    .trim()
    .replace(/^(?:pp:)?slide:/i, '')
    .slice(0, 64);
  return {
    error:
      `Shape ${JSON.stringify(key.slice(0, 64))} was not found on that slide (there are no shape ` +
      `numbers or names). For the slide title use title as the shape: shape pp:shape:${slideKey}:title ` +
      `"…" (or shape=title on /format-shape). Otherwise use one of its shape ids: ` +
      `${shown || 'the slide has no shapes'}. A shape you added has the id in its add-shape result.`,
    reason: 'shape_not_found',
  };
}

/**
 * `id 86 (TextBox, title: "What is Tokenomics?"), id 85 (GeometricShape, no text)` — for a
 * not-found shape error. Ids and types alone left the model guessing: after "Shape 1" failed on a
 * deck whose first shape was a decorative line, it wrote the title into the line. Shape text is
 * document content: quoted, single-line and clipped, like the read listing. Read-only.
 */
/**
 * A text box on a slide whose title/body placeholders are still empty lands ON TOP of them: the
 * "Click to add title / text" prompts stay behind it and the slide reads as unstructured (UAT
 * 2026-10-16, a "daftar isi" written as one text box over a blank Title and Content slide). Refuse
 * it before anything is written and name the placeholders to write into instead. Read-only; needs
 * placeholder types (PowerPointApi 1.8), else no check.
 */
async function emptyPlaceholderMessage(
  ctx: PowerPoint.RequestContext,
  slideId: string,
  slideRef: string,
): Promise<string | undefined> {
  if (!isSet('PowerPointApi', '1.8')) return undefined;
  const shapes = ctx.presentation.slides.getItem(slideId).shapes;
  shapes.load('items/id,items/type');
  await ctx.sync();
  const placeholders = shapes.items
    .filter((shape) => shape.type === 'Placeholder')
    .map((shape) => {
      const format = shape.placeholderFormat;
      format.load('type');
      const range = shape.textFrame.textRange;
      range.load('text');
      return { id: shape.id, format, range };
    });
  if (placeholders.length === 0) return undefined;
  try {
    await ctx.sync();
  } catch {
    return undefined; // A placeholder without text (picture, chart): no check rather than a block.
  }
  const empty = placeholders.filter(({ range }) => (range.text ?? '').trim() === '');
  const title = empty.find(({ format }) => TITLE_PLACEHOLDERS.includes(format.type));
  const body = empty.find(({ format }) => BODY_PLACEHOLDERS.includes(format.type));
  if (!title && !body) return undefined;
  const slide = slideRef.trim().replace(/^(?:pp:)?slide:/i, '');
  const targets = [
    ...(title ? [`the title: shape pp:shape:${slide}:title "…"`] : []),
    ...(body
      ? [`the body (one line per item): shape pp:shape:${slide}:${body.id} "line 1\\nline 2"`]
      : []),
  ];
  return (
    'This slide has empty layout placeholders; a text box would sit on top of them. Write into ' +
    `them instead — ${targets.join('; ')}. Not applied.`
  );
}

async function describeShapes(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<string> {
  let texts = new Map<PowerPoint.Shape, string>();
  let title: PowerPoint.Shape | undefined;
  try {
    const readable = await withoutNonTextPlaceholders(ctx, shapes.filter(isReadableShape));
    texts = await readShapeContent(ctx, readable);
    title = await findTitleShape(ctx, shapes);
  } catch {
    // Fall back to ids and types.
  }
  return shapes
    .map((shape) => {
      const text = (texts.get(shape) ?? '').replace(/\s+/g, ' ').trim();
      const clipped = text.length > 40 ? `${text.slice(0, 40)}…` : text;
      const label = shape === title ? 'title: ' : '';
      return `id ${shape.id} (${shape.type}, ${text ? `${label}${JSON.stringify(clipped)}` : 'no text'})`;
    })
    .join(', ');
}

/**
 * A slide's title shape: its Title/CenterTitle placeholder, else the text shape set in the largest
 * type (see {@link largestTextShape}), else the first shape with text.
 */
async function findTitleShape(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<PowerPoint.Shape | undefined> {
  const { titleShape } = await findTextPlaceholders(ctx, shapes);
  if (titleShape && isSet('PowerPointApi', '1.8')) return titleShape;
  const readable = await withoutNonTextPlaceholders(ctx, shapes.filter(isReadableShape));
  const texts = await readShapeContent(ctx, readable);
  const withText = readable.filter(
    (shape) => shape.type !== 'Table' && (texts.get(shape) ?? '').trim(),
  );
  if (withText.length <= 1) return withText[0];
  return (await largestTextShape(ctx, withText, texts)) ?? withText[0];
}

/**
 * The shape whose first visible character is set in the largest font (ties: the topmost). Decks
 * without title placeholders (Google Slides exports, hand-built slides) often put a small label
 * ("SECTION / 01") above the title, so the first shape with text was the label, not the title.
 * The first character, because a whole range with mixed runs reports `font.size` as null. One
 * read-only sync; undefined when the host reports no sizes.
 */
async function largestTextShape(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
  texts: Map<PowerPoint.Shape, string>,
): Promise<PowerPoint.Shape | undefined> {
  try {
    const probes = shapes.map((shape) => {
      const text = texts.get(shape) ?? '';
      const first = shape.textFrame.textRange.getSubstring(Math.max(0, text.search(/\S/)), 1);
      first.font.load('size');
      shape.load('top');
      return [shape, first] as const;
    });
    await ctx.sync();
    let best: { shape: PowerPoint.Shape; size: number; top: number } | undefined;
    for (const [shape, first] of probes) {
      const size = first.font.size ?? 0;
      const top = shape.top ?? Number.POSITIVE_INFINITY;
      if (size > 0 && (!best || size > best.size || (size === best.size && top < best.top))) {
        best = { shape, size, top };
      }
    }
    return best?.shape;
  } catch {
    return undefined;
  }
}

/** The deck's slide size when the host reports it (PowerPointApi 1.10), else undefined. */
async function knownSlideSize(
  ctx: PowerPoint.RequestContext,
): Promise<{ width: number; height: number } | undefined> {
  if (!isSet('PowerPointApi', '1.10')) return undefined;
  const size = await slideSize(ctx);
  return size === FALLBACK_SLIDE_SIZE ? undefined : size;
}

/**
 * `add-shape` needs explicit placement (the capability's contract: "never infer placement from
 * prose"). Without it the host drops the shape at the top-left corner — so "near the bottom" came
 * out at the top. The message gives the slide size so the next turn can place it.
 */
function missingGeometryMessage(
  size: { width: number; height: number } | undefined,
  geometry: { left?: number; top?: number; width?: number; height?: number },
): string | undefined {
  const missing = (['left', 'top', 'width', 'height'] as const).filter(
    (k) => geometry[k] === undefined,
  );
  if (missing.length === 0) return undefined;
  const slide = size
    ? `The slide is ${size.width} × ${size.height} pt (the bottom edge is top=${size.height}). `
    : '';
  return (
    `add-shape needs an explicit position: missing ${missing.join(', ')}. ${slide}` +
    'Give left= top= width= height= in points, e.g. near the bottom: ' +
    (size
      ? `left=${Math.round(size.width * 0.1)} top=${Math.round(size.height * 0.82)} width=${Math.round(size.width * 0.6)} height=${Math.round(size.height * 0.1)}.`
      : 'left=72 top=330 width=400 height=40.')
  );
}

/**
 * A corrective message when explicit geometry would place the shape outside the slide (models
 * guess a 960 × 540 slide on a 720 × 405 deck), so the next turn can fix it; undefined when it fits
 * or the size is unknown.
 */
function offSlideMessage(
  size: { width: number; height: number } | undefined,
  geometry: { left?: number; top?: number; width?: number; height?: number },
): string | undefined {
  if (!size) return undefined;
  const { left = 0, top = 0, width = 0, height = 0 } = geometry;
  if (left >= 0 && top >= 0 && left + width <= size.width && top + height <= size.height) {
    return undefined;
  }
  return (
    `That position is outside the slide, which is ${size.width} × ${size.height} pt: keep ` +
    `left + width ≤ ${size.width} and top + height ≤ ${size.height} ` +
    `(got left ${left}, top ${top}, width ${width}, height ${height}).`
  );
}

/** `" (PowerPoint: <code>)"` for an Office.js error, so an "uncertain" result says why. */
function hostErrorSuffix(error: unknown): string {
  const code = (error as { code?: unknown } | undefined)?.code;
  return typeof code === 'string' && /^[\w.-]{1,64}$/.test(code) ? ` (PowerPoint: ${code})` : '';
}

/** Shapes listed per slide in the outline; the rest stay reachable through `read slide:N`. */
const MAX_SUMMARY_SHAPES = 8;

/**
 * `shapes: pp:shape:256:2 Placeholder: SLA Terms; pp:shape:256:5 Image`. The renderer escapes and clips the whole line;
 * shape text (document content) also loses the characters this listing and the inventory use as
 * structure, so it cannot fake another shape, close the summary or mimic an entry.
 */
function shapeSummary(slideKey: string, shapes: ShapeElement[]): string | undefined {
  if (shapes.length === 0) return undefined;
  const listed = shapes.slice(0, MAX_SUMMARY_SHAPES).map((shape) => {
    const text = Array.from(
      shape.text
        .replace(/[;,:()[\]=]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim(),
    );
    const clipped = text.length > 24 ? `${text.slice(0, 24).join('')}…` : text.join('');
    // The full ref: a bare id made the model write `shape 2 "…"`, which needs the slide too.
    const label = [`pp:shape:${slideKey}:${shape.shapeId}`, shape.type].filter(Boolean).join(' ');
    return clipped ? `${label}: ${clipped}` : label;
  });
  if (shapes.length > MAX_SUMMARY_SHAPES)
    listed.push(`+${shapes.length - MAX_SUMMARY_SHAPES} more`);
  return `shapes: ${listed.join('; ')}`;
}

/** The id of the slide at `index`, read in a fresh request (undefined if it can't be read). */
async function readNewSlideIdAt(index: number, ids: NewSlideIds): Promise<string | undefined> {
  try {
    return await PowerPoint.run(async (ctx) => {
      // Only the new slide's id: if the deck shifted since, keep the id the write verified.
      const check = queueNewSlideCheck(ctx, index, ids);
      await ctx.sync();
      return check();
    });
  } catch {
    return undefined;
  }
}

function slideNotFound(
  req: ActuationRequest,
  message: string,
  reason?: ActuationErrorReason,
): ActuationResult {
  return {
    ok: false,
    changeId: req.changeId,
    kind: req.kind,
    degraded: true,
    error: { code: 'target_conflict', message, ...(reason ? { reason } : {}) },
  };
}

/** Actual dispatch keys; conformance checks these against the advertised capabilities. */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = PowerPointBridge.handledActuations;
