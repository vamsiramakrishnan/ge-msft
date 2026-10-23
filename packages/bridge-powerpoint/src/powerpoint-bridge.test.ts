import { afterEach, describe, it, expect, vi } from 'vitest';
import {
  asChangeId,
  DocStateSnapshotSchema,
  ResolvedContextSchema,
  type ActuationRequest,
  type ContextRef,
} from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { renderDocState } from '@ge/content';
import { MAX_READ_SLIDES, PowerPointBridge, tableHeight } from './powerpoint-bridge.js';

/**
 * Behavioural tests for the {@link PowerPointBridge} host wiring — the file that touches the
 * `PowerPoint.run` / `Office` globals directly. There is no host-port indirection here (unlike the
 * Word bridge), so we install a self-contained in-memory PowerPoint + Office simulator onto
 * `globalThis` for the duration of each test. The simulator re-implements ONLY the slice of the
 * Office.js object model the bridge actually drives (slides → shapes → textRange text, the
 * compose/insert WRITE paths, the requirement gate, and the Office event bus). The REAL bridge runs
 * unchanged against it, so these assert real semantics: capability gating, bounded reads, slide
 * addressing, the two insert-slide write paths, provenance/location stamping, and event wiring.
 */

/* ─────────────────────────── in-memory host model ─────────────────────────── */

interface ShapeSeed {
  text: string;
  /** Simulated `Shape.type`; defaults to 'TextBox'. */
  type?: string;
  /** Simulated `placeholderFormat.containedType` for a 'Placeholder' shape (null = text/empty). */
  containedType?: string | null;
  /** Simulated `placeholderFormat.type` for a 'Placeholder' shape (e.g. 'Title', 'Body'). */
  placeholderType?: string;
  /** The host refuses this shape's text frame even though its type suggests one. */
  refuseText?: boolean;
  /** Simulated ShapeFill.foregroundColor (undefined until first set). */
  fillColor?: string;
  /** Simulated ShapeLineFormat.color (undefined until first set). */
  lineColor?: string;
  /** Simulated ShapeFont fields on the shape's text range. */
  font?: {
    bold?: boolean | null;
    italic?: boolean | null;
    underline?: string | null;
    color?: string | null;
    size?: number | null;
    name?: string | null;
  };
  /** Simulated native table grid backing `Shape.getTable()` (table shapes only). */
  tableGrid?: string[][];
  /** Recorded setZOrder positions applied to this shape. */
  zOrderCalls?: string[];
  /** Simulated `Shape.top` in points. */
  top?: number;
  /** Simulated `Shape.left` / `.width` / `.height` in points. */
  left?: number;
  width?: number;
  height?: number;
  /** Set when the bridge sets `textFrame.autoSizeSetting`. */
  autoSize?: string;
  /** Set when the bridge hides the paragraphs' bullets (`bulletFormat.visible = false`). */
  bulletsVisible?: boolean;
  /** Font size of the text's first character (`textRange.getSubstring(i, 1).font.size`). */
  firstCharSize?: number;
}
interface SlideSeed {
  id: string;
  shapes: ShapeSeed[];
  /**
   * Set on a just-added slide: its `id` is PROVISIONAL until the current `PowerPoint.run` ends, when
   * it becomes this settled id. Like PowerPoint for the web, writes addressed through the
   * provisional id (`getItem(id)` or the `items` list) fail the next sync; `getItemAt` works.
   */
  settledId?: string;
  /**
   * With {@link DeckSeed.settleIdsAfterSyncs}: syncs left before the provisional id settles inside
   * the same request, as PowerPoint for the web does live (2026-10-16).
   */
  settleAfterSyncs?: number;
  /** With {@link DeckSeed.layoutShapesOnSettle}: the layout's shapes, shown once the id settles. */
  pendingShapes?: ShapeSeed[];
}

/** A no-op write check, for fakes not reached through a slide. */
const NO_GUARD = (): void => undefined;
interface DeckSeed {
  slides: SlideSeed[];
  /** Zero-based indices of selected slides (the bridge reads items[0]). */
  selectedIndices: number[];
  selectedShapeIds: string[];
  insertedDecks: string[];
  insertedDeckOptions: PowerPoint.InsertSlideOptions[];
  /** Shapes a `slides.add()` slide gets from the default layout (default: title + body placeholders). */
  newSlideShapes?: ShapeSeed[];
  /**
   * The slide master's layouts (`SlideLayout.type`, PowerPointApi 1.8). `slides.add({ layoutId })`
   * with one of these ids gives the slide its shapes; an unknown id is ignored and the default
   * layout applies, as PowerPoint for the web does (live 2026-10-16).
   */
  layouts?: Array<{ id: string; type?: string; name?: string; shapes: ShapeSeed[] }>;
  /**
   * A slide added WITH a layout has no shapes until its id settles — PowerPoint for the web, live
   * 2026-10-16 (about a second after the add).
   */
  layoutShapesOnSettle?: boolean;
  /** Runs at the start of every sync, before ids settle: a co-author's change mid-request. */
  onSync?: () => void;
  /** The options each `slides.add()` was called with. */
  addOptions?: Array<PowerPoint.AddSlideOptions | undefined>;
  /** Make `Slide.moveTo` fail its sync, like a host that rejects the move. */
  failMoveTo?: boolean;
  /**
   * Settle a just-added slide's id after this many syncs instead of at the end of the request.
   * Live PowerPoint for the web (2026-10-16): `add` + sync, then `items/id` + sync still reads the
   * provisional id, and the next sync already reads the settled one — so 3.
   */
  settleIdsAfterSyncs?: number;
  addedShapes: Array<{
    slideId: string;
    kind: 'textBox' | 'geometric' | 'line' | 'table';
    detail: string;
    options: PowerPoint.ShapeAddOptions | PowerPoint.TableAddOptions | undefined;
  }>;
}

/** A cell whose text reads/writes one slot of the backing {@link ShapeSeed.tableGrid}. */
class FakeTableCell {
  constructor(
    private readonly grid: string[][],
    private readonly rowIndex: number,
    private readonly columnIndex: number,
  ) {}
  get text(): string {
    return this.grid[this.rowIndex]?.[this.columnIndex] ?? '';
  }
  set text(value: string) {
    const row = this.grid[this.rowIndex];
    if (row && this.columnIndex < row.length) row[this.columnIndex] = value;
  }
}

/** A simulated native table whose storage is the owning shape's `tableGrid` seed. */
class FakeTable {
  constructor(private readonly shape: ShapeSeed) {}
  load(_p?: string): this {
    return this;
  }
  get values(): string[][] {
    return (this.shape.tableGrid ?? []).map((row) => [...row]);
  }
  getCellOrNullObject(rowIndex: number, columnIndex: number): FakeTableCell {
    const grid = (this.shape.tableGrid ??= []);
    while (grid.length <= rowIndex) grid.push([]);
    const row = grid[rowIndex];
    if (row) while (row.length <= columnIndex) row.push('');
    return new FakeTableCell(grid, rowIndex, columnIndex);
  }
}

class FakeFont {
  constructor(
    private readonly font: NonNullable<ShapeSeed['font']>,
    private readonly guard: () => void = NO_GUARD,
  ) {}
  load(_p?: string): this {
    return this;
  }
  get bold(): boolean | null {
    return this.font.bold ?? null;
  }
  set bold(v: boolean) {
    this.guard();
    this.font.bold = v;
  }
  get italic(): boolean | null {
    return this.font.italic ?? null;
  }
  set italic(v: boolean) {
    this.guard();
    this.font.italic = v;
  }
  get underline(): string | null {
    return this.font.underline ?? null;
  }
  set underline(v: 'None' | 'Single') {
    this.guard();
    this.font.underline = v;
  }
  get color(): string | null {
    return this.font.color ?? null;
  }
  set color(v: string) {
    this.guard();
    this.font.color = v;
  }
  get size(): number | null {
    return this.font.size ?? null;
  }
  set size(v: number) {
    this.guard();
    this.font.size = v;
  }
  get name(): string | null {
    return this.font.name ?? null;
  }
  set name(v: string) {
    this.guard();
    this.font.name = v;
  }
}

class FakeTextRange {
  readonly font: FakeFont;
  constructor(
    private readonly shape: ShapeSeed,
    private readonly guard: () => void = NO_GUARD,
  ) {
    this.shape.font ??= {};
    this.font = new FakeFont(this.shape.font, guard);
  }
  get text(): string {
    return this.shape.text;
  }
  set text(v: string) {
    this.guard();
    this.shape.text = v;
  }
  load(_p?: string): this {
    return this;
  }
  /** `paragraphFormat.bulletFormat.visible` (PowerPointApi 1.4), recorded on the seed. */
  get paragraphFormat(): { bulletFormat: { visible: boolean } } {
    const shape = this.shape;
    return {
      bulletFormat: {
        get visible(): boolean {
          return shape.bulletsVisible ?? true;
        },
        set visible(v: boolean) {
          shape.bulletsVisible = v;
        },
      },
    };
  }
  getSubstring(start: number, length: number): { font: { size: number | null; load(): void } } {
    // Like the host: a substring outside the text fails the next sync.
    if (start + length > this.shape.text.length) pendingHostError = 'InvalidArgument: substring';
    return { font: { size: this.shape.firstCharSize ?? null, load: () => undefined } };
  }
}

class FakeShapeFill {
  constructor(private readonly shape: ShapeSeed) {}
  load(_p?: string): this {
    return this;
  }
  get foregroundColor(): string {
    return this.shape.fillColor ?? '';
  }
  set foregroundColor(v: string) {
    this.shape.fillColor = v;
  }
}

class FakeShapeLineFormat {
  constructor(private readonly shape: ShapeSeed) {}
  load(_p?: string): this {
    return this;
  }
  get color(): string {
    return this.shape.lineColor ?? '';
  }
  set color(v: string) {
    this.shape.lineColor = v;
  }
}

/** Shape types the simulated host gives a text frame (mirrors the host's InvalidArgument throw). */
const FAKE_TEXT_SHAPE_TYPES = new Set([
  'GeometricShape',
  'TextBox',
  'Placeholder',
  'Callout',
  'Freeform',
]);

/**
 * Like the real host, touching a text frame / placeholder format the shape doesn't have queues an
 * operation that fails at the NEXT `ctx.sync()` — failing the whole batch, not just that shape.
 */
let pendingHostError: string | undefined;

class FakeShape {
  private readonly frame: { textRange: FakeTextRange };
  readonly fill: FakeShapeFill;
  readonly lineFormat: FakeShapeLineFormat;
  constructor(
    readonly shape: ShapeSeed,
    readonly id: string,
    readonly isNullObject = false,
    guard: () => void = NO_GUARD,
  ) {
    this.frame = { textRange: new FakeTextRange(shape, guard) };
    this.fill = new FakeShapeFill(shape);
    this.lineFormat = new FakeShapeLineFormat(shape);
  }
  get type(): string {
    return this.shape.type ?? 'TextBox';
  }
  get top(): number | undefined {
    return this.shape.top;
  }
  get left(): number {
    return this.shape.left ?? 0;
  }
  get width(): number {
    return this.shape.width ?? 0;
  }
  get height(): number {
    return this.shape.height ?? 0;
  }
  set height(value: number) {
    this.shape.height = value;
  }
  get textFrame(): { textRange: FakeTextRange; autoSizeSetting?: string } {
    const contained = this.shape.containedType;
    const refuses =
      this.shape.refuseText === true ||
      !FAKE_TEXT_SHAPE_TYPES.has(this.type) ||
      (this.type === 'Placeholder' && contained != null && !FAKE_TEXT_SHAPE_TYPES.has(contained));
    if (refuses) pendingHostError = `InvalidArgument: ${this.type} ${this.id} has no TextFrame`;
    const shape = this.shape;
    return {
      textRange: this.frame.textRange,
      set autoSizeSetting(value: string) {
        shape.autoSize = value;
      },
    };
  }
  get placeholderFormat(): {
    containedType: string | null;
    type: string;
    load(_p?: string): unknown;
  } {
    if (this.type !== 'Placeholder') {
      pendingHostError = `GeneralException: ${this.id} is not a placeholder`;
    }
    const format = {
      containedType: this.shape.containedType ?? null,
      type: this.shape.placeholderType ?? 'Unsupported',
      load: () => format,
    };
    return format;
  }
  load(_p?: string): this {
    return this;
  }
  setZOrder(position: string): void {
    (this.shape.zOrderCalls ??= []).push(position);
  }
  getTable(): FakeTable {
    if (this.type !== 'Table') pendingHostError = `InvalidArgument: ${this.id} is not a table`;
    return new FakeTable(this.shape);
  }
}

class FakeShapeCollection {
  private built: FakeShape[] = [];
  constructor(
    private readonly slide: SlideSeed,
    private readonly deck: DeckSeed,
    private readonly slideId: string,
    private readonly guard: () => void = NO_GUARD,
  ) {}
  /** Read live, like loaded Office.js values after a sync: a settling slide's shapes can appear. */
  get items(): FakeShape[] {
    if (this.built.length !== this.slide.shapes.length) {
      this.built = this.slide.shapes.map(
        (s, i) => new FakeShape(s, `${this.slide.id}-shape-${i}`, false, this.guard),
      );
    }
    return this.built;
  }
  load(_p?: string): this {
    return this;
  }
  getItemAt(index: number): FakeShape {
    const shape = this.items[index];
    if (!shape) throw new Error(`fake-powerpoint: no shape at ${index}`);
    return shape;
  }
  getItemOrNullObject(id: string): FakeShape {
    const index = this.items.findIndex((shape) => shape.id === id);
    const existing = this.items[index];
    if (existing) return existing;
    return new FakeShape(
      { text: '', zOrderCalls: [] },
      id || `${this.slideId}-missing-shape`,
      true,
    );
  }

  /** Append a simulated shape to the host model and the live collection, then return it. */
  private append(seed: ShapeSeed): FakeShape {
    const items = this.items;
    const id = `${this.slideId}-shape-${items.length}`;
    this.slide.shapes.push(seed);
    this.guard();
    const shape = new FakeShape(seed, id, false, this.guard);
    items.push(shape);
    return shape;
  }

  addTextBox(text: string, options?: PowerPoint.ShapeAddOptions): FakeShape {
    this.deck.addedShapes.push({
      slideId: this.slideId,
      kind: 'textBox',
      detail: text,
      options,
    });
    return this.append({ text, zOrderCalls: [] });
  }

  addGeometricShape(geometryType: string, options?: PowerPoint.ShapeAddOptions): FakeShape {
    this.deck.addedShapes.push({
      slideId: this.slideId,
      kind: 'geometric',
      detail: geometryType,
      options,
    });
    return this.append({ text: '', zOrderCalls: [] });
  }

  addLine(connectorType?: string, options?: PowerPoint.ShapeAddOptions): FakeShape {
    this.deck.addedShapes.push({
      slideId: this.slideId,
      kind: 'line',
      detail: connectorType ?? 'Straight',
      options,
    });
    return this.append({ text: '', zOrderCalls: [] });
  }

  addTable(rowCount: number, columnCount: number, options?: PowerPoint.TableAddOptions): FakeShape {
    this.deck.addedShapes.push({
      slideId: this.slideId,
      kind: 'table',
      detail: `${rowCount}x${columnCount}`,
      options,
    });
    const grid: string[][] = Array.from({ length: rowCount }, () =>
      Array.from({ length: columnCount }, () => ''),
    );
    return this.append({ text: '', type: 'Table', tableGrid: grid, zOrderCalls: [] });
  }
}

class FakeSlide {
  constructor(
    private readonly slide: SlideSeed,
    private readonly position: number,
    private readonly seed: DeckSeed,
    /**
     * Reached via `getItemAt` (by position) rather than by id or the `items` list. Like Office.js,
     * LOADING the proxy re-addresses it by id (`slides.getItem(id)`), so `load()` clears this.
     */
    private byPosition = false,
  ) {}
  /** A write through this slide: fails the next sync if it addresses a provisional id. */
  private readonly guard = (): void => {
    if (this.slide.settledId !== undefined && !this.byPosition) {
      pendingHostError = 'InvalidParam passed to GetItem(id)';
    }
  };
  get id(): string {
    return this.slide.id;
  }
  get index(): number {
    requireSlideIndex('index');
    return this.position;
  }
  load(props?: string): this {
    requireSlideIndex(props);
    this.byPosition = false;
    return this;
  }
  get shapes(): FakeShapeCollection {
    return new FakeShapeCollection(this.slide, this.seed, this.slide.id, this.guard);
  }
  /** `Slide.slideMaster` (PowerPointApi 1.3): every fake slide shares the one master. */
  get slideMaster(): FakeSlideMaster {
    return new FakeSlideMaster(this.seed);
  }
  /** `Slide.moveTo` (PowerPointApi 1.8): reorders the deck. */
  moveTo(slideIndex: number): void {
    if (this.seed.failMoveTo) {
      pendingHostError = 'GeneralException';
      return;
    }
    const from = this.seed.slides.indexOf(this.slide);
    this.seed.slides.splice(from, 1);
    this.seed.slides.splice(slideIndex, 0, this.slide);
  }
  /** `Slide.delete` (PowerPointApi 1.2). */
  delete(): void {
    this.guard();
    const at = this.seed.slides.indexOf(this.slide);
    if (at >= 0) this.seed.slides.splice(at, 1);
  }
  setSelectedShapes(shapeIds: string[]): void {
    this.seed.selectedIndices = [this.index];
    this.seed.selectedShapeIds = [...shapeIds];
  }
}

/** A slide master whose layouts come from {@link DeckSeed.layouts}. */
class FakeSlideMaster {
  constructor(private readonly seed: DeckSeed) {}
  get layouts(): FakeLayoutCollection {
    return new FakeLayoutCollection(this.seed);
  }
}

class FakeLayoutCollection {
  constructor(private readonly seed: DeckSeed) {}
  get items(): Array<{ id: string; type?: string; name: string }> {
    return (this.seed.layouts ?? []).map(({ id, type, name }) => ({ id, type, name: name ?? '' }));
  }
  load(_p?: string): this {
    return this;
  }
  getItemAt(index: number): {
    id: string;
    type?: string;
    name: string;
    shapes: FakeShapeCollection;
    load(_p?: string): unknown;
  } {
    const layout = this.seed.layouts?.[index];
    if (!layout) throw new Error(`fake-powerpoint: no layout at ${index}`);
    return {
      id: layout.id,
      type: layout.type,
      name: layout.name ?? '',
      shapes: new FakeShapeCollection(
        { id: layout.id, shapes: layout.shapes.map((shape) => ({ ...shape, zOrderCalls: [] })) },
        this.seed,
        layout.id,
      ),
      load: () => undefined,
    };
  }
}

class FakeSlideCollection {
  items: FakeSlide[];
  /** Recorded calls so tests can assert what the bridge drove. */
  addCalls = 0;
  constructor(private readonly seed: DeckSeed) {
    this.items = seed.slides.map((s, i) => new FakeSlide(s, i, seed));
  }
  load(props?: string): this {
    requireSlideIndex(props);
    return this;
  }
  getCount(): { value: number } {
    return { value: this.seed.slides.length };
  }
  add(options?: PowerPoint.AddSlideOptions): void {
    this.addCalls += 1;
    (this.seed.addOptions ??= []).push(options);
    const layout = this.seed.layouts?.find((l) => l.id === options?.layoutId);
    const late = layout && this.seed.layoutShapesOnSettle;
    const n = this.seed.slides.length + 1;
    this.seed.slides.push({
      id: `provisional-${n}#0`,
      settledId: `sim-slide-${n}`,
      ...(late
        ? { pendingShapes: layout.shapes.map((shape) => ({ ...shape, zOrderCalls: [] })) }
        : {}),
      ...(this.seed.settleIdsAfterSyncs === undefined
        ? {}
        : { settleAfterSyncs: this.seed.settleIdsAfterSyncs }),
      shapes: (
        layout?.shapes ??
        this.seed.newSlideShapes ?? [
          { text: '', type: 'Placeholder', placeholderType: 'Title', zOrderCalls: [] },
          { text: '', type: 'Placeholder', placeholderType: 'Body', zOrderCalls: [] },
        ]
      )
        .filter(() => !late)
        .map((shape) => ({ ...shape, zOrderCalls: [] })),
    });
    this.items = this.seed.slides.map((s, i) => new FakeSlide(s, i, this.seed));
  }
  getItemAt(index: number): FakeSlide {
    const slide = this.seed.slides[index];
    if (!slide) throw new Error(`fake-powerpoint: no slide at ${index}`);
    return new FakeSlide(slide, index, this.seed, true);
  }
  getItem(id: string): FakeSlide {
    const index = this.seed.slides.findIndex((s) => s.id === id);
    const slide = this.seed.slides[index];
    if (!slide) throw new Error(`fake-powerpoint: no slide ${id}`);
    return new FakeSlide(slide, index, this.seed);
  }
}

/** Model the real property version, not only the method's requirement set. */
function requireSlideIndex(props?: string): void {
  if (
    /\bindex\b/.test(props ?? '') &&
    !Office.context.requirements.isSetSupported('PowerPointApi', '1.8')
  )
    throw new Error('Slide.index requires PowerPointApi 1.8');
}

class FakePresentation {
  insertCalls: string[] = [];
  /** Simulated `PageSetup` (PowerPointApi 1.10): a 10" × 5.625" deck, in points. */
  readonly pageSetup = {
    slideWidth: 720,
    slideHeight: 405,
    load(_p?: string): unknown {
      return this;
    },
  };
  constructor(private readonly seed: DeckSeed) {}
  get slides(): FakeSlideCollection {
    return new FakeSlideCollection(this.seed);
  }
  get slideMasters(): { getItemAt(index: number): FakeSlideMaster } {
    return { getItemAt: () => new FakeSlideMaster(this.seed) };
  }
  getSelectedSlides(): FakeSlideCollection {
    const all = new FakeSlideCollection(this.seed);
    const sel = new FakeSlideCollection(this.seed);
    sel.items = this.seed.selectedIndices
      .map((i) => all.items[i])
      .filter((s): s is FakeSlide => s !== undefined);
    return sel;
  }
  setSelectedSlides(slideIds: string[]): void {
    this.seed.selectedIndices = slideIds
      .map((id) => this.seed.slides.findIndex((s) => s.id === id))
      .filter((i) => i >= 0);
  }
  insertSlidesFromBase64(b64: string, options?: PowerPoint.InsertSlideOptions): void {
    this.insertCalls.push(b64);
    this.seed.insertedDecks.push(b64);
    this.seed.insertedDeckOptions.push(options ?? {});
    this.seed.slides.push({ id: `sim-inserted-${this.seed.insertedDecks.length}`, shapes: [] });
  }
}

/** A just-added slide gets its real id, and any layout shapes the host was still building. */
function settle(slide: SlideSeed): void {
  if (slide.settledId !== undefined) slide.id = slide.settledId;
  if (slide.pendingShapes) slide.shapes = slide.pendingShapes;
  delete slide.settledId;
  delete slide.settleAfterSyncs;
  delete slide.pendingShapes;
}

class FakeContext {
  readonly presentation: FakePresentation;
  syncs = 0;
  constructor(private readonly seed: DeckSeed) {
    this.presentation = new FakePresentation(seed);
  }
  sync(): Promise<void> {
    this.syncs += 1;
    this.seed.onSync?.();
    for (const slide of this.seed.slides) {
      if (slide.settleAfterSyncs === undefined || slide.settledId === undefined) continue;
      slide.settleAfterSyncs -= 1;
      if (slide.settleAfterSyncs <= 0) settle(slide);
    }
    const error = pendingHostError;
    pendingHostError = undefined;
    return error ? Promise.reject(new Error(error)) : Promise.resolve();
  }
}

type OfficeHandler = () => void;
interface OfficeRegistry {
  added: Array<{ type: unknown; handler: OfficeHandler }>;
  removed: Array<{ type: unknown; handler: OfficeHandler }>;
}

interface InstallOpts {
  /** Requirement-set support: name → max supported version (numeric compare on the version). */
  requirements?: Record<string, number>;
  /** When true, every Office addHandlerAsync throws (host without the event bus). */
  officeThrowsOnAdd?: boolean;
  /** When true, removeHandlerAsync throws (teardown best-effort path). */
  officeThrowsOnRemove?: boolean;
}

interface Installed {
  ctxRef: { ctx?: FakeContext };
  office: OfficeRegistry;
  restore(): void;
}

/**
 * Install fake `PowerPoint` + `Office` globals modelling `seed`, returning hooks to inspect the
 * last context and the Office handler registry. `requirements` drives `isSetSupported(name, ver)`.
 */
function install(seed: DeckSeed, opts: InstallOpts = {}): Installed {
  const reqs = opts.requirements ?? { PowerPointApi: 8 };
  const ctxRef: { ctx?: FakeContext } = {};
  const office: OfficeRegistry = { added: [], removed: [] };

  pendingHostError = undefined;
  const prevPP = (globalThis as Record<string, unknown>).PowerPoint;
  const prevOffice = (globalThis as Record<string, unknown>).Office;

  (globalThis as Record<string, unknown>).PowerPoint = {
    run: async <T>(cb: (ctx: FakeContext) => Promise<T>): Promise<T> => {
      const ctx = new FakeContext(seed);
      ctxRef.ctx = ctx;
      try {
        return await cb(ctx);
      } finally {
        // A request completed: just-added slides get their real ids (see SlideSeed.settledId).
        for (const slide of seed.slides) {
          if (slide.settledId !== undefined) settle(slide);
        }
      }
    },
  };

  (globalThis as Record<string, unknown>).Office = {
    EventType: {
      DocumentSelectionChanged: 'DocumentSelectionChanged',
      ActiveViewChanged: 'ActiveViewChanged',
    },
    context: {
      requirements: {
        isSetSupported: (name: string, version?: string): boolean => {
          const max = reqs[name];
          if (max === undefined) return false;
          // `max` 1.x = highest supported minor version; >= 2 = every 1.x set. Compare minors as
          // integers so '1.10' is newer than '1.5' (a float compare would read it as 1.1).
          if (max >= 2) return true;
          const minor = Number((version ?? '1.0').split('.')[1] ?? '0');
          return minor <= Math.round((max - 1) * 10);
        },
      },
      document: {
        addHandlerAsync: (type: unknown, handler: OfficeHandler): void => {
          if (opts.officeThrowsOnAdd) throw new Error('no event bus');
          office.added.push({ type, handler });
        },
        removeHandlerAsync: (type: unknown, arg: { handler: OfficeHandler }): void => {
          if (opts.officeThrowsOnRemove) throw new Error('cannot remove');
          office.removed.push({ type, handler: arg.handler });
        },
      },
    },
  };

  return {
    ctxRef,
    office,
    restore: () => {
      (globalThis as Record<string, unknown>).PowerPoint = prevPP;
      (globalThis as Record<string, unknown>).Office = prevOffice;
    },
  };
}

let installed: Installed | undefined;
afterEach(() => {
  installed?.restore();
  installed = undefined;
});

function deck(slides: SlideSeed[], selectedIndices: number[] = []): DeckSeed {
  // Deep-copy so shared fixtures (SAMPLE_SLIDES) aren't mutated by write paths across tests.
  const copy = slides.map((s) => ({
    id: s.id,
    shapes: s.shapes.map((sh) => ({
      ...sh,
      font: sh.font ? { ...sh.font } : undefined,
      tableGrid: sh.tableGrid?.map((row) => [...row]),
      zOrderCalls: [...(sh.zOrderCalls ?? [])],
    })),
  }));
  return {
    slides: copy,
    selectedIndices,
    selectedShapeIds: [],
    insertedDecks: [],
    insertedDeckOptions: [],
    addedShapes: [],
  };
}

function insertSlide(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'insert-slide', surface: 'powerpoint', params };
}

function setShapeText(
  slideId: string,
  shapeId: string,
  text: string,
  id = 'shape-1',
): ActuationRequest {
  return {
    changeId: asChangeId(id),
    kind: 'set-shape-text',
    surface: 'powerpoint',
    params: { target: { slideId, shapeId }, text },
  };
}

const SAMPLE_SLIDES: SlideSeed[] = [
  {
    id: 's1',
    shapes: [
      { text: 'SLA Terms', zOrderCalls: [] },
      { text: '99.5% contracted\nMonthly window', zOrderCalls: [] },
    ],
  },
  {
    id: 's2',
    shapes: [
      { text: 'Roster', zOrderCalls: [] },
      { text: 'Pat, Sam', zOrderCalls: [] },
    ],
  },
  {
    id: 's3',
    shapes: [
      { text: 'Risk Summary', zOrderCalls: [] },
      { text: 'SLA gap flagged', zOrderCalls: [] },
    ],
  },
];

const WHOLE_DECK: ContextRef = {
  id: 'pp:deck',
  kind: 'document',
  surface: 'powerpoint',
  title: 'Whole deck',
};

function contextText(ctx: Array<{ value: { as: string; text?: string } }>): string {
  return ctx.map((c) => (c.value.as === 'text' ? (c.value.text ?? '') : '')).join('\n');
}

/* ───────────────────────────── getCapabilities ───────────────────────────── */

describe('PowerPointBridge surface + capabilities', () => {
  it('reports the powerpoint surface and a capability manifest advertising insert-slide', () => {
    const bridge = new PowerPointBridge();
    expect(bridge.surface).toBe('powerpoint');
    const caps = bridge.getCapabilities();
    expect(caps.surface).toBe('powerpoint');
    expect(caps.actuations.map((a) => a.kind)).toContain('insert-slide');
  });
});

describe('PowerPoint read compatibility', () => {
  it.each([1.4, 1.5, 1.7, 1.8])(
    'keeps correct slide positions across read ports on API %s',
    async (version) => {
      installed = install(deck(SAMPLE_SLIDES, [1]), { requirements: { PowerPointApi: version } });
      const bridge = new PowerPointBridge();
      const snapshot = await bridge.captureDocState();
      expect(snapshot?.inventory.map((entry) => entry.id)).toEqual([
        'slide:s1',
        'slide:s2',
        'slide:s3',
      ]);
      const slide = {
        id: 'pp:slide:s2',
        kind: 'slide',
        surface: 'powerpoint',
        title: 'Slide 2',
      } as const;
      expect(contextText(await bridge.resolveContext(slide))).toContain('Pat, Sam');
      expect(contextText(await bridge.readRange('slide:2'))).toContain('Pat, Sam');
      expect(contextText(await bridge.searchDocument('Roster'))).toContain('Pat, Sam');
      const shape = await bridge.resolveContext({
        id: 'pp:shape:s2:s2-shape-1',
        kind: 'shape',
        surface: 'powerpoint',
        title: 'Shape on slide 2',
      });
      expect(contextText(shape)).toContain('Pat, Sam');
      expect(shape[0]?.ref.title).toBe('Shape s2-shape-1 on slide 2');
      const refs = await bridge.listContext();
      const selected = await bridge.resolveContext({
        id: 'selection',
        kind: 'selection',
        surface: 'powerpoint',
        title: 'Selected slide',
      });
      if (version >= 1.5) {
        expect(refs[0]?.title).toBe('Slide 2: Roster');
        expect(contextText(selected)).toContain('Pat, Sam');
        expect(snapshot?.selection?.title).toBe('Slide 2 (slide:s2)');
      } else {
        expect(refs.map((ref) => ref.id)).toEqual(['pp:deck']);
        expect(selected).toEqual([]);
        expect(snapshot?.selection).toBeUndefined();
      }
    },
  );
});

/* ───────────────────────────── listContext ───────────────────────────── */

describe('PowerPointBridge.listContext', () => {
  it('degrades to a deck-only ref when getSelectedSlides (1.5) is unsupported', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.4 } });
    const refs = await new PowerPointBridge().listContext();
    expect(refs).toEqual([
      { id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' },
    ]);
  });

  it('lists the selected slide (live, 1-based title) plus the whole deck when 1.5 is supported', async () => {
    installed = install(deck(SAMPLE_SLIDES, [1]));
    const refs = await new PowerPointBridge().listContext();
    expect(refs).toHaveLength(4);
    expect(refs[0]).toEqual({
      id: 'pp:slide:s2',
      kind: 'slide',
      surface: 'powerpoint',
      title: 'Slide 2: Roster', // index 1 → human "Slide 2"
      preview: 'Roster',
      live: true,
      anchor: { matchText: 'Roster', locator: 'slide:s2' },
      hostRef: { type: 'powerpoint.slide', slideId: 's2' },
    });
    expect(refs[1]).toMatchObject({
      id: 'pp:shape:s2:s2-shape-0',
      kind: 'shape',
      title: 'Shape s2-shape-0 on slide 2',
      preview: 'Roster',
    });
    expect(refs[2]).toMatchObject({
      id: 'pp:shape:s2:s2-shape-1',
      kind: 'shape',
      title: 'Shape s2-shape-1 on slide 2',
      preview: 'Pat, Sam',
    });
    expect(refs[3]).toMatchObject({ id: 'pp:deck', kind: 'document' });
  });

  it('lists only shapes with text, titled by shape id (not position)', async () => {
    installed = install(
      deck(
        [
          {
            id: 's1',
            shapes: [
              { text: '', type: 'GeometricShape', zOrderCalls: [] }, // a decorative line
              { text: 'What is Tokenomics?', zOrderCalls: [] },
            ],
          },
        ],
        [0],
      ),
    );
    const shapes = (await new PowerPointBridge().listContext()).filter((r) => r.kind === 'shape');
    expect(shapes.map((r) => [r.id, r.title])).toEqual([
      ['pp:shape:s1:s1-shape-1', 'Shape s1-shape-1 on slide 1'],
    ]);
  });

  it('lists only the deck when nothing is selected', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const refs = await new PowerPointBridge().listContext();
    expect(refs).toEqual([
      { id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' },
    ]);
  });
});

/* ───────────────────────────── resolveContext ───────────────────────────── */

describe('PowerPointBridge.resolveContext', () => {
  const slideRef: ContextRef = {
    id: 'pp:slide:s2',
    kind: 'slide',
    surface: 'powerpoint',
    title: 'Slide 2',
  };

  it('resolves a slide ref by host slide id independent of the current selection', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const ctx = await new PowerPointBridge().resolveContext(slideRef);
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s2')).toBe(true);
  });

  it('resolves an exact slide ref when no slide is selected', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const ctx = await new PowerPointBridge().resolveContext(slideRef);
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s2')).toBe(true);
  });

  it('returns [] for a stale slide ref', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    expect(
      await new PowerPointBridge().resolveContext({
        id: 'pp:slide:missing',
        kind: 'slide',
        surface: 'powerpoint',
        title: 'Missing slide',
      }),
    ).toEqual([]);
  });

  it('resolves a shape ref to the exact shape text', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const ctx = await new PowerPointBridge().resolveContext({
      id: 'pp:shape:s2:s2-shape-1',
      kind: 'shape',
      surface: 'powerpoint',
      title: 'Roster body',
    });
    expect(ctx).toHaveLength(1);
    expect(ctx[0]?.ref.id).toBe('pp:shape:s2:s2-shape-1');
    expect(ctx[0]?.value.as).toBe('text');
    if (ctx[0]?.value.as === 'text') expect(ctx[0].value.text).toBe('Pat, Sam');
  });

  it('treats selection and shape refs the same way as slide refs', async () => {
    installed = install(deck(SAMPLE_SLIDES, [2]));
    const sel = await new PowerPointBridge().resolveContext({
      id: 'x',
      kind: 'selection',
      surface: 'powerpoint',
      title: 'Selection',
    });
    expect(sel.some((c) => c.ref.anchor?.locator === 'slide:s3')).toBe(true);
  });

  it('resolves a document ref by reading every slide in the deck, anchored per slide', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const ctx = await new PowerPointBridge().resolveContext({
      id: 'pp:deck',
      kind: 'document',
      surface: 'powerpoint',
      title: 'Whole deck',
    });
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const locators = ctx.map((c) => c.ref.anchor?.locator);
    expect(locators).toContain('slide:s1');
    expect(locators).toContain('slide:s2');
    expect(locators).toContain('slide:s3');
  });

  it('reads the whole deck in a constant number of syncs, whatever the slide count', async () => {
    const wholeDeck: ContextRef = {
      id: 'pp:deck',
      kind: 'document',
      surface: 'powerpoint',
      title: 'Whole deck',
    };
    const slides = (count: number): SlideSeed[] =>
      Array.from({ length: count }, (_, i) => ({
        id: `s${i}`,
        shapes: [
          { text: `Title ${i}`, zOrderCalls: [] },
          { text: `Body ${i}`, zOrderCalls: [] },
        ],
      }));

    installed = install(deck(slides(3)));
    await new PowerPointBridge().resolveContext(wholeDeck);
    const small = installed.ctxRef.ctx?.syncs;
    installed.restore();

    installed = install(deck(slides(30)));
    const ctx = await new PowerPointBridge().resolveContext(wholeDeck);
    expect(installed.ctxRef.ctx?.syncs).toBe(small);
    expect(small).toBeLessThanOrEqual(3);
    expect(ctx.map((c) => c.value.as === 'text' && c.value.text).join('\n')).toContain('Body 29');
  });

  it('skips shapes without a text frame instead of failing the whole-deck read', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Deck title', zOrderCalls: [] },
            { text: '', type: 'Image', zOrderCalls: [] },
            { text: '', type: 'Table', zOrderCalls: [] },
          ],
        },
        { id: 's2', shapes: [{ text: 'Second slide', zOrderCalls: [] }] },
      ]),
    );
    const ctx = await new PowerPointBridge().resolveContext({
      id: 'pp:deck',
      kind: 'document',
      surface: 'powerpoint',
      title: 'Whole deck',
    });
    const text = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(text).toContain('Deck title');
    expect(text).toContain('Second slide');
  });

  it('skips a placeholder holding a picture without falling back to per-shape reads', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Offsite', type: 'Placeholder', containedType: null, zOrderCalls: [] },
            { text: '', type: 'Placeholder', containedType: 'Image', zOrderCalls: [] },
            { text: 'Lisbon, 8-10 October', type: 'Placeholder', zOrderCalls: [] },
          ],
        },
        { id: 's2', shapes: [{ text: 'Next slide', zOrderCalls: [] }] },
      ]),
    );
    const ctx = await new PowerPointBridge().resolveContext(WHOLE_DECK);
    const text = contextText(ctx);
    expect(text).toContain('Offsite');
    expect(text).toContain('Lisbon, 8-10 October');
    expect(text).toContain('Next slide');
    // slides + shapes + placeholder formats + text: the batch path, no per-shape retries.
    expect(installed.ctxRef.ctx?.syncs).toBe(4);
  });

  it('falls back to per-shape reads when the host refuses an unexpected shape', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Kept title', zOrderCalls: [] },
            { text: 'secret', type: 'GeometricShape', refuseText: true, zOrderCalls: [] },
            { text: 'Kept body', zOrderCalls: [] },
          ],
        },
      ]),
    );
    const text = contextText(await new PowerPointBridge().resolveContext(WHOLE_DECK));
    expect(text).toContain('Kept title');
    expect(text).toContain('Kept body');
    expect(text).not.toContain('secret');
  });

  it('still reads the deck on a host without placeholderFormat (PowerPointApi < 1.8)', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Old host title', type: 'Placeholder', zOrderCalls: [] },
            { text: '', type: 'Placeholder', containedType: 'Image', zOrderCalls: [] },
          ],
        },
      ]),
      { requirements: { PowerPointApi: 1.5 } },
    );
    const text = contextText(await new PowerPointBridge().resolveContext(WHOLE_DECK));
    expect(text).toContain('Old host title');
  });

  it('reads native table cells into the whole deck, in the same round-trip as text', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Wave 3 site budgets', zOrderCalls: [] },
            {
              text: '',
              type: 'Table',
              tableGrid: [
                ['Site', 'Budget'],
                ['Madrid', '$310K'],
                ['', ''],
              ],
              zOrderCalls: [],
            },
          ],
        },
      ]),
    );
    const text = contextText(await new PowerPointBridge().resolveContext(WHOLE_DECK));
    expect(text).toContain('Site | Budget');
    expect(text).toContain('Madrid | $310K');
    // slides + shapes + content (text and tables together): no extra sync for tables.
    expect(installed.ctxRef.ctx?.syncs).toBe(3);
  });

  it('skips tables on a host without Table.values (PowerPointApi < 1.8)', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Budgets', zOrderCalls: [] },
            { text: '', type: 'Table', tableGrid: [['Madrid', '$310K']], zOrderCalls: [] },
          ],
        },
      ]),
      { requirements: { PowerPointApi: 1.5 } },
    );
    const text = contextText(await new PowerPointBridge().resolveContext(WHOLE_DECK));
    expect(text).toContain('Budgets');
    expect(text).not.toContain('Madrid');
  });

  it('resolves a table shape chip to its cells and a picture shape chip to nothing', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: '', type: 'Image', zOrderCalls: [] },
            {
              text: '',
              type: 'Table',
              tableGrid: [
                ['Aisle', 'Picks / hr'],
                ['C-14', '312'],
              ],
              zOrderCalls: [],
            },
          ],
        },
      ]),
    );
    const shapeRef = (shapeId: string): ContextRef => ({
      id: `pp:shape:s1:${shapeId}`,
      kind: 'shape',
      surface: 'powerpoint',
      title: 'Shape',
      hostRef: { type: 'powerpoint.shape', slideId: 's1', shapeId },
    });
    const bridge = new PowerPointBridge();
    expect(await bridge.resolveContext(shapeRef('s1-shape-0'))).toEqual([]);
    expect(contextText(await bridge.resolveContext(shapeRef('s1-shape-1')))).toContain(
      'C-14 | 312',
    );
  });

  it('bounds the whole-deck read to MAX_READ_SLIDES slides', async () => {
    installed = install(
      deck(
        Array.from({ length: MAX_READ_SLIDES + 5 }, (_, i) => ({
          id: `s${i}`,
          shapes: [{ text: `Title ${i + 1}`, zOrderCalls: [] }],
        })),
      ),
    );
    const ctx = await new PowerPointBridge().resolveContext({
      id: 'pp:deck',
      kind: 'document',
      surface: 'powerpoint',
      title: 'Whole deck',
    });
    const text = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(text).toContain(`Title ${MAX_READ_SLIDES}`);
    expect(text).not.toContain(`Title ${MAX_READ_SLIDES + 1}`);
  });
});

/* ───────────────────────────── captureDocState ───────────────────────────── */

describe('PowerPointBridge.captureDocState', () => {
  it.each([
    [[0], 'Slide 1 (slide:s1)'],
    [[1], 'Slide 2 (slide:s2)'],
    [[0, 2], 'Slide 1 (slide:s1), Slide 3 (slide:s3)'],
  ] as const)('identifies the live selected slides %s in the snapshot', async (indices, title) => {
    installed = install(deck(SAMPLE_SLIDES, [...indices]));
    const snapshot = await new PowerPointBridge().captureDocState();
    expect(snapshot?.selection).toEqual({ kind: 'slide', title });
  });

  it('omits selection when the host lacks selected-slide support or nothing is selected', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.4 } });
    expect((await new PowerPointBridge().captureDocState())?.selection).toBeUndefined();
    installed.restore();
    installed = install(deck(SAMPLE_SLIDES, []));
    expect((await new PowerPointBridge().captureDocState())?.selection).toBeUndefined();
  });

  it('keeps the deck snapshot if the optional selected-slide read fails', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const run = PowerPoint.run;
    const host = vi
      .spyOn(PowerPoint, 'run')
      .mockImplementationOnce(run)
      .mockRejectedValueOnce(new Error('selection temporarily unavailable'));
    try {
      const snapshot = await new PowerPointBridge().captureDocState();
      expect(snapshot?.inventory.map((entry) => entry.id)).toEqual([
        'slide:s1',
        'slide:s2',
        'slide:s3',
      ]);
      expect(snapshot?.selection).toBeUndefined();
    } finally {
      host.mockRestore();
    }
  });

  it('returns undefined when TextRange.text (1.4) is unsupported', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.3 } });
    expect(await new PowerPointBridge().captureDocState()).toBeUndefined();
  });

  it('returns undefined for an empty deck', async () => {
    installed = install(deck([], []));
    expect(await new PowerPointBridge().captureDocState()).toBeUndefined();
  });

  it('builds a valid snapshot whose inventory comes from the slides and bumps version each capture', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const bridge = new PowerPointBridge();

    const first = await bridge.captureDocState();
    expect(first).toBeDefined();
    if (!first) return;
    expect(() => DocStateSnapshotSchema.parse(first)).not.toThrow();
    expect(first.surface).toBe('powerpoint');
    expect(first.version).toBe(1);
    expect(first.outline.some((o) => o.text.includes('SLA Terms'))).toBe(true);

    const second = await bridge.captureDocState();
    expect(second?.version).toBe(2);
  });

  it('lists the last slide of a long deck by number (live 2026-10-09)', async () => {
    // One inventory entry per bullet filled the cap by about slide 11 of 15, hiding a new slide.
    const fifteen: SlideSeed[] = Array.from({ length: 15 }, (_, i) => ({
      id: `p${i}`,
      shapes: [{ text: `Topic ${i + 1}` }, { text: 'one\ntwo\nthree\nfour' }],
    }));
    installed = install(deck(fifteen, [0]));
    const snap = await new PowerPointBridge().captureDocState();
    expect(snap?.truncated).toBeUndefined();
    expect(snap?.inventory).toHaveLength(15);
    expect(snap?.inventory.at(-1)).toEqual({
      kind: 'slide',
      id: 'slide:p14',
      title: 'Topic 15',
      summary:
        'slide 15; shapes: pp:shape:p14:p14-shape-0: Topic 15; pp:shape:p14:p14-shape-1: one two three four',
    });
  });

  it("lists each slide's shape ids so shape commands need no guessing", async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const snap = await new PowerPointBridge().captureDocState();
    const rendered = snap ? renderDocState(snap) : '';
    const line = rendered.split('\n').find((l) => l.includes('SLA Terms') && l.includes('shapes:'));
    expect(line).toBeDefined();
    expect(line).toContain(
      '(slide 1; shapes: pp:shape:s1:s1-shape-0: SLA Terms; pp:shape:s1:s1-shape-1: 99.5% contracted',
    );
    // Once per slide, not once per body line that shares the slide's locator.
    expect(rendered.match(/shapes: pp:shape:s1:s1-shape-0/g)).toHaveLength(1);
  });

  it('strips listing structure from shape text so a slide cannot forge shapes or entries', async () => {
    installed = install(
      deck(
        [
          {
            id: 'h1',
            shapes: [{ text: 'Plan' }, { text: 'x) [slide] ref=9; id 7 Placeholder: T' }],
          },
        ],
        [0],
      ),
    );
    const snap = await new PowerPointBridge().captureDocState();
    const line = (snap ? renderDocState(snap) : '').split('\n').find((l) => l.includes('shapes:'));
    expect(line).toBeDefined();
    const shapeText = line!.slice(line!.indexOf('pp:shape:h1:h1-shape-1'), line!.lastIndexOf(')'));
    expect(shapeText).not.toMatch(/[;()[\]=]/);
    expect(line!.match(/; pp:shape:/g)).toHaveLength(1);
  });

  it('reads at most MAX_READ_SLIDES slides (bounded per-turn read, ADR-0006)', async () => {
    // Give every slide a unique title; the snapshot must not list beyond the cap.
    const many: SlideSeed[] = Array.from({ length: MAX_READ_SLIDES + 10 }, (_, i) => ({
      id: `m${i}`,
      shapes: [{ text: `Title ${i}` }],
    }));
    installed = install(deck(many, [0]));
    const snap = await new PowerPointBridge().captureDocState();
    expect(snap).toBeDefined();
    if (!snap) return;
    // The slide past the cap must not appear in the inventory.
    expect(snap.outline.some((o) => o.text.includes(`Title ${MAX_READ_SLIDES + 5}`))).toBe(false);
    expect(snap.outline.some((o) => o.text.includes('Title 0'))).toBe(true);
  });
});

/* ───────────────────────────── searchDocument ───────────────────────────── */

describe('PowerPointBridge.searchDocument', () => {
  it('returns [] for an empty/whitespace query without touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    expect(await new PowerPointBridge().searchDocument('   ')).toEqual([]);
    // No PowerPoint.run was invoked for an empty query.
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('returns [] when the host predates TextRange.text (1.4)', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.3 } });
    expect(await new PowerPointBridge().searchDocument('SLA')).toEqual([]);
  });

  it('returns the slides whose title/body match the query (case-insensitive)', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const ctx = await new PowerPointBridge().searchDocument('sla');
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const locators = ctx.map((c) => c.ref.anchor?.locator);
    expect(locators).toContain('slide:s1'); // "SLA Terms"
    expect(locators).toContain('slide:s3'); // "SLA gap flagged"
    expect(locators).not.toContain('slide:s2'); // Roster — no SLA
  });

  it('returns [] when nothing matches', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    expect(await new PowerPointBridge().searchDocument('nonexistent-token')).toEqual([]);
  });
});

/* ───────────────────────────── readRange ───────────────────────────── */

describe('PowerPointBridge.readRange', () => {
  it('resolves an addressable slide selector (1-based) to that single slide', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const ctx = await new PowerPointBridge().readRange('slide:1');
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s1')).toBe(true);
    // Only one slide resolved, not the whole deck.
    expect(new Set(ctx.map((c) => c.ref.anchor?.locator))).toEqual(new Set(['slide:s1']));
  });

  it('accepts a bare 1-based index and maps to the zero-based slide', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const ctx = await new PowerPointBridge().readRange('3');
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s3')).toBe(true);
  });

  it('returns [] for an unaddressable selector without invoking the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    expect(await new PowerPointBridge().readRange('Agenda')).toEqual([]);
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('returns [] for an out-of-range slide index (degrade, not crash)', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    expect(await new PowerPointBridge().readRange('slide:99')).toEqual([]);
  });

  it('accepts a host slide id and the pp:slide: ref form', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    for (const selector of ['s2', 'slide:s2', 'pp:slide:s2', 'pp:slide:2']) {
      const ctx = await new PowerPointBridge().readRange(selector);
      expect(new Set(ctx.map((c) => c.ref.anchor?.locator))).toEqual(new Set(['slide:s2']));
    }
    const last = await new PowerPointBridge().readRange('last');
    expect(last.some((c) => c.ref.anchor?.locator === 'slide:s3')).toBe(true);
  });

  it('lists the slide shapes with ids, types and the title, so shape commands can address them', async () => {
    installed = install(
      deck(
        [
          {
            id: 's1',
            shapes: [
              { text: 'FY26 Plan', type: 'Placeholder', placeholderType: 'Title', zOrderCalls: [] },
              { text: 'Grow ARR', type: 'TextBox', zOrderCalls: [] },
              { text: '', type: 'Image', zOrderCalls: [] },
            ],
          },
        ],
        [],
      ),
    );
    const ctx = await new PowerPointBridge().readRange('slide:1');
    const text = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(text).toContain(
      'for shape commands use slide=1 and shape=<id>; shape=title is id s1-shape-0',
    );
    expect(text).toContain('- id s1-shape-0 · Placeholder · title · "FY26 Plan"');
    expect(text).toContain('- id s1-shape-1 · TextBox · "Grow ARR"');
    expect(text).toContain('- id s1-shape-2 · Image');
  });

  it('returns [] when the host predates TextRange.text (1.4)', async () => {
    installed = install(deck(SAMPLE_SLIDES, []), { requirements: { PowerPointApi: 1.3 } });
    expect(await new PowerPointBridge().readRange('slide:1')).toEqual([]);
  });
});

/* ───────────────────────────── revealContext ───────────────────────────── */

describe('PowerPointBridge.revealContext', () => {
  it('selects an addressable slide ref by host slide id', async () => {
    const seed = deck(SAMPLE_SLIDES, [0]);
    installed = install(seed);
    const bridge = new PowerPointBridge();
    const ref: ContextRef = {
      id: 'pp:slide:s3',
      kind: 'slide',
      surface: 'powerpoint',
      title: 'Slide 3',
    };

    expect(bridge.canRevealContext(ref)).toBe(true);
    await bridge.revealContext(ref);

    expect(seed.selectedIndices).toEqual([2]);
  });

  it('selects a shape when the ref carries slide and shape ids', async () => {
    const seed = deck(SAMPLE_SLIDES, [0]);
    installed = install(seed);
    await new PowerPointBridge().revealContext({
      id: 'pp:shape:s2:s2-shape-1',
      kind: 'shape',
      surface: 'powerpoint',
      title: 'Body text',
    });

    expect(seed.selectedIndices).toEqual([1]);
    expect(seed.selectedShapeIds).toEqual(['s2-shape-1']);
  });

  it('does not advertise deck-wide refs as revealable', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    expect(
      new PowerPointBridge().canRevealContext({
        id: 'pp:deck',
        kind: 'document',
        surface: 'powerpoint',
        title: 'Whole deck',
      }),
    ).toBe(false);
  });
});

/* ───────────────────────────── actuate: dispatch + insert-slide ───────────────────────────── */

describe('PowerPointBridge.actuate dispatch', () => {
  it('returns an unsupported error for a kind PowerPoint cannot handle, echoing the changeId', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('chg-x'),
      kind: 'write-cells',
      surface: 'powerpoint',
      params: { cells: [['1']] },
    });
    expect(res).toMatchObject({
      ok: false,
      changeId: asChangeId('chg-x'),
      kind: 'write-cells',
      error: { code: 'unsupported' },
    });
    // No host write happened for an unsupported kind.
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

describe('PowerPointBridge.actuate insert-slide (empty)', () => {
  it('rejects an insert with no base64, no title, and no bullets before touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: '', bullets: [] } }),
    );
    expect(res).toMatchObject({
      ok: false,
      kind: 'insert-slide',
      error: { code: 'empty_slide' },
    });
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('rejects an insert whose only content is whitespace notes (notes do not satisfy the gate)', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: '', bullets: [], notes: '   ' } }),
    );
    expect(res.ok).toBe(false);
    expect(res.error?.code).toBe('empty_slide');
  });
});

describe('PowerPointBridge.actuate insert-slide (base64 prebuilt deck)', () => {
  it('merges a prebuilt base64 deck via insertSlidesFromBase64 and reports inserted-deck', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ ooxml: 'UEsDBBQ=' }, 'chg-deck'),
    );
    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-deck'),
      kind: 'insert-slide',
      location: 'inserted-deck',
    });
    // The base64 payload reached the host merge path exactly once.
    expect(d.insertedDecks).toEqual(['UEsDBBQ=']);
    expect(d.insertedDeckOptions).toEqual([{ targetSlideId: 's3' }]);
  });

  it('imports an explicit generated deck artifact in one call with formatting and slide count', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide(
        {
          deck: {
            format: 'pptx',
            base64: 'compiled-deck',
            slideCount: 2,
            formatting: 'UseDestinationTheme',
            targetSlideId: 's1',
            specFingerprint: '9f34a012',
          },
        },
        'chg-generated-deck',
      ),
    );
    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-generated-deck'),
      kind: 'insert-slide',
      location: 'inserted-deck:2',
    });
    expect(d.insertedDecks).toEqual(['compiled-deck']);
    expect(d.insertedDeckOptions).toEqual([
      { formatting: 'UseDestinationTheme', targetSlideId: 's1' },
    ]);
  });

  it('prefers the base64 path over native compose when both are present (no slides.add)', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const before = d.slides.length;
    const res = await new PowerPointBridge().actuate(
      insertSlide({ ooxml: 'UEs=', slide: { title: 'Ignored', bullets: ['x'] } }),
    );
    expect(res.location).toBe('inserted-deck');
    expect(d.insertedDecks).toEqual(['UEs=']);
    // Native compose would have seeded a 2-shape placeholder slide; base64 seeds a 0-shape marker.
    const appended = d.slides[d.slides.length - 1];
    expect(d.slides.length).toBe(before + 1);
    expect(appended?.shapes).toHaveLength(0);
  });
});

describe('PowerPointBridge.actuate insert-slide (native compose)', () => {
  it('appends a slide and writes the title into shape[0] and bullets into shape[1]', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const before = d.slides.length;
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'New Topic', bullets: ['Point A', 'Point B'] } }, 'chg-native'),
    );
    expect(res.ok).toBe(true);
    expect(res.changeId).toBe(asChangeId('chg-native'));
    expect(res.location).toBe('slide:sim-slide-4'); // appended at index `before`
    // Exactly one slide was appended.
    expect(d.slides.length).toBe(before + 1);
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes[0]?.text).toBe('New Topic');
    expect(appended?.shapes[1]?.text).toBe('Point A\rPoint B');
  });

  it('turns a literal \\n in slide text into a real line break (UAT 2026-10-16)', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    await new PowerPointBridge().actuate(
      insertSlide({
        slide: { title: 'Lifecycle\\nControl', bullets: ['Pre-Deployment\\n• Checklist'] },
      }),
    );
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes[0]?.text).toBe('Lifecycle\nControl');
    expect(appended?.shapes[1]?.text).toBe('Pre-Deployment\n• Checklist');
  });

  it('adds a new slide with the Title and Content layout, not the default cover (UAT #12/#13)', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    const ph = (placeholderType: string): ShapeSeed => ({
      text: '',
      type: 'Placeholder',
      placeholderType,
      zOrderCalls: [],
    });
    d.newSlideShapes = [ph('CenterTitle'), ph('Subtitle')];
    d.layouts = [
      { id: 'L-cover', name: 'Title Slide', shapes: [ph('CenterTitle'), ph('Subtitle')] },
      { id: 'L-two', name: 'Two Content', shapes: [ph('Title'), ph('Body'), ph('Body')] },
      { id: 'L-content', name: 'Title and Content', shapes: [ph('Title'), ph('Body')] },
    ];
    installed = install(d);
    await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Daftar Isi', bullets: ['Latar belakang', 'Monitoring'] } }),
    );
    // The layout comes from the last slide's master, so no slideMasterId (PR #5); bullets are
    // paragraphs (\r), not line breaks.
    expect(d.addOptions?.at(-1)).toEqual({ layoutId: 'L-content' });
    const added = d.slides.at(-1)!;
    expect(added.shapes[0]?.text).toBe('Daftar Isi');
    expect(added.shapes[1]?.text).toBe('Latar belakang\rMonitoring');
  });

  it('never writes bullets into a cover subtitle: they go in a left-aligned text box', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    d.newSlideShapes = [
      { text: '', type: 'Placeholder', placeholderType: 'CenterTitle', zOrderCalls: [] },
      { text: '', type: 'Placeholder', placeholderType: 'Subtitle', zOrderCalls: [] },
    ];
    installed = install(d);
    await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Roadmap', bullets: ['Q4 2026', 'Q1 2027'] } }),
    );
    const added = d.slides.at(-1)!;
    expect(added.shapes[0]?.text).toBe('Roadmap');
    expect(added.shapes[1]?.text).toBe('');
    expect(added.shapes.at(-1)?.text).toBe('• Q4 2026\r• Q1 2027');
  });

  it('writes a title-only slide without touching the body placeholder', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    installed = install(d);
    await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Just A Title', bullets: [] } }),
    );
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes[0]?.text).toBe('Just A Title');
    // Body placeholder stays empty when there are no bullets.
    expect(appended?.shapes[1]?.text).toBe('');
  });

  it('writes bullets even when the title is empty (body-only compose)', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: '', bullets: ['Only a bullet'] } }),
    );
    expect(res.ok).toBe(true);
    const appended = d.slides[d.slides.length - 1];
    // Title placeholder untouched (empty); bullets written to the body placeholder.
    expect(appended?.shapes[0]?.text).toBe('');
    expect(appended?.shapes[1]?.text).toBe('Only a bullet');
  });
});

describe('PowerPointBridge.actuate insert-slide (at a position)', () => {
  it('moves the new slide to the requested position after writing its text', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const before = d.slides.map((s) => s.id);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Agenda', bullets: ['Goals'] }, target: { slideIndex: 1 } }),
    );
    expect(res.ok).toBe(true);
    expect(res.error).toBeUndefined();
    expect(d.slides[1]?.shapes[0]?.text).toBe('Agenda');
    expect(d.slides[1]?.shapes[1]?.text).toBe('Goals');
    expect(d.slides.filter((s) => before.includes(s.id)).map((s) => s.id)).toEqual(before);
    expect(d.slides[0]?.id).toBe(before[0]);
  });

  it('appends when the position is past the end', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Last', bullets: [] }, target: { slideIndex: 99 } }),
    );
    expect(res.ok).toBe(true);
    expect(d.slides.at(-1)?.shapes[0]?.text).toBe('Last');
  });

  it('refuses cleanly, writing nothing, on a host without Slide.moveTo', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d, { requirements: { PowerPointApi: 1.5 } });
    const before = d.slides.length;
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Agenda', bullets: [] }, target: { slideIndex: 1 } }),
    );
    expect(res.ok).toBe(false);
    expect(res.error?.code).toBe('unsupported');
    expect(d.slides.length).toBe(before);
  });

  it('reports an unconfirmed move as landed-but-uncertain, never as plain success', async () => {
    const d = { ...deck(SAMPLE_SLIDES, [0]), failMoveTo: true };
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Agenda', bullets: [] }, target: { slideIndex: 1 } }),
    );
    expect(res.ok).toBe(true);
    expect(res.error?.code).toBe('not_moved');
    expect(d.slides.at(-1)?.shapes[0]?.text).toBe('Agenda');
  });
});

describe('PowerPointBridge.actuate insert-slide (layout without placeholders)', () => {
  it('matches placeholders by type, not position', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [
      { text: '', type: 'Placeholder', placeholderType: 'Body' },
      { text: '', type: 'Placeholder', placeholderType: 'Title' },
    ];
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Typed', bullets: ['One', 'Two'] } }),
    );
    expect(res.ok).toBe(true);
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes[1]?.text).toBe('Typed');
    expect(appended?.shapes[0]?.text).toBe('One\rTwo');
    expect(d.addedShapes).toEqual([]);
  });

  it('adds sized text boxes when the default layout has no placeholders (Blank layout)', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [];
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'this is a test slide', bullets: ['First', 'Second'] } }),
    );
    expect(res.ok).toBe(true);
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes.map((shape) => shape.text)).toEqual([
      'this is a test slide',
      '• First\r• Second',
    ]);
    // Both boxes sit inside the 720 × 405 pt slide the host reported.
    for (const added of d.addedShapes) {
      const box = added.options as PowerPoint.ShapeAddOptions;
      expect(added.kind).toBe('textBox');
      expect((box.left ?? 0) + (box.width ?? 0)).toBeLessThanOrEqual(720);
      expect((box.top ?? 0) + (box.height ?? 0)).toBeLessThanOrEqual(405);
    }
    expect(appended?.shapes[0]?.font?.bold).toBe(true);
  });

  it('adds only a body text box when the layout has a title placeholder but no body', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [{ text: '', type: 'Placeholder', placeholderType: 'Title' }];
    installed = install(d);
    await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Heading', bullets: ['Point'] } }),
    );
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes.map((shape) => shape.text)).toEqual(['Heading', '• Point']);
    expect(d.addedShapes.map((added) => added.kind)).toEqual(['textBox']);
  });

  it('falls back to the first two placeholders in order below PowerPointApi 1.8', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [
      { text: '', type: 'Placeholder' },
      { text: '', type: 'Placeholder' },
    ];
    installed = install(d, { requirements: { PowerPointApi: 1.5 } });
    await new PowerPointBridge().actuate(insertSlide({ slide: { title: 'T', bullets: ['B'] } }));
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes.map((shape) => shape.text)).toEqual(['T', 'B']);
  });

  it('writes to the just-added slide by position and reports its settled id (PowerPoint web)', async () => {
    // The fake gives a new slide a provisional id until the request ends, and fails writes addressed
    // through that id — as PowerPoint for the web does (5010 "InvalidParam passed to GetItem(id)").
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [];
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Q4 plan', bullets: ['hire 5 engineers'] } }),
    );
    expect(res).toMatchObject({ ok: true, location: 'slide:sim-slide-2' });
    expect(d.slides[1]?.shapes.map((shape) => shape.text)).toEqual([
      'Q4 plan',
      '• hire 5 engineers',
    ]);
  });

  it('writes the new slide when its id settles mid-request (PowerPoint web, live 2026-10-16)', async () => {
    // Live: the id read right after `add` is provisional and the very next sync already reads the
    // settled one. That is the same slide, not a co-author's change — it must not fail the insert.
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.settleIdsAfterSyncs = 3;
    installed = install(d);
    const bridge = new PowerPointBridge();
    const first = await bridge.actuate(
      insertSlide({ slide: { title: 'Latar belakang', bullets: ['Adopsi AI meningkat'] } }),
    );
    const second = await bridge.actuate(
      insertSlide({ slide: { title: 'Risiko utama', bullets: ['Bias model'] } }),
    );
    expect(first).toMatchObject({ ok: true, location: 'slide:sim-slide-2' });
    expect(second).toMatchObject({ ok: true, location: 'slide:sim-slide-3' });
    expect(d.slides.map((slide) => slide.shapes.map((shape) => shape.text))).toEqual([
      ['Existing'],
      ['Latar belakang', 'Adopsi AI meningkat'],
      ['Risiko utama', 'Bias model'],
    ]);
  });

  describe('new-slide layout (UAT 2026-10-16: bullets landed in a centred Title Slide subtitle)', () => {
    const LAYOUTS = [
      {
        id: 'L-title',
        type: 'Title',
        shapes: [
          { text: '', type: 'Placeholder', placeholderType: 'CenterTitle' },
          { text: '', type: 'Placeholder', placeholderType: 'Subtitle' },
        ],
      },
      {
        id: 'L-content',
        type: 'Object',
        shapes: [
          { text: '', type: 'Placeholder', placeholderType: 'Title' },
          { text: '', type: 'Placeholder', placeholderType: 'Body' },
        ],
      },
      {
        id: 'L-title-only',
        type: 'TitleOnly',
        shapes: [{ text: '', type: 'Placeholder', placeholderType: 'Title' }],
      },
    ] satisfies DeckSeed['layouts'];
    const layoutDeck = (): DeckSeed => {
      const d = deck([{ id: 's1', shapes: [{ text: 'Agenda' }] }], [0]);
      d.layouts = structuredClone(LAYOUTS);
      return d;
    };

    it('puts a slide with bullets on the Title and Content layout, not the first layout', async () => {
      const d = layoutDeck();
      installed = install(d);
      const res = await new PowerPointBridge().actuate(
        insertSlide({
          slide: { title: 'Risiko utama', bullets: ['Bias model', 'Kebocoran data'] },
        }),
      );
      expect(res.ok).toBe(true);
      expect(d.addOptions).toEqual([{ layoutId: 'L-content' }]);
      expect(d.slides[1]?.shapes.map((shape) => [shape.placeholderType, shape.text])).toEqual([
        ['Title', 'Risiko utama'],
        ['Body', 'Bias model\rKebocoran data'],
      ]);
    });

    it('waits for a layout slide to settle so its text lands in the placeholders (PowerPoint web)', async () => {
      const d = layoutDeck();
      d.settleIdsAfterSyncs = 6;
      d.layoutShapesOnSettle = true;
      installed = install(d);
      const res = await new PowerPointBridge().actuate(
        insertSlide({ slide: { title: 'Monitoring', bullets: ['Drift model', 'Audit berkala'] } }),
      );
      expect(res).toMatchObject({ ok: true, location: 'slide:sim-slide-2' });
      expect(d.slides[1]?.shapes.map((shape) => [shape.placeholderType, shape.text])).toEqual([
        ['Title', 'Monitoring'],
        ['Body', 'Drift model\rAudit berkala'],
      ]);
      expect(d.addedShapes).toEqual([]);
    });

    it('hides the layout bullet when the slide lines number themselves', async () => {
      const d = layoutDeck();
      installed = install(d);
      await new PowerPointBridge().actuate(
        insertSlide({
          slide: { title: 'Daftar Isi', bullets: ['1. Latar belakang', '2. Roadmap'] },
        }),
      );
      expect(d.slides[1]?.shapes[1]?.bulletsVisible).toBe(false);
    });

    it('puts a title-only slide on the Title Only layout', async () => {
      const d = layoutDeck();
      installed = install(d);
      await new PowerPointBridge().actuate(insertSlide({ slide: { title: 'Q&A', bullets: [] } }));
      expect(d.addOptions).toEqual([{ layoutId: 'L-title-only' }]);
    });

    it('keeps the host default when the deck has no content layout', async () => {
      const d = layoutDeck();
      d.layouts = d.layouts?.filter((layout) => layout.type === 'Title');
      installed = install(d);
      const res = await new PowerPointBridge().actuate(
        insertSlide({ slide: { title: 'T', bullets: ['B'] } }),
      );
      expect(res.ok).toBe(true);
      expect(d.addOptions).toEqual([undefined]);
    });

    it('keeps the host default without layout types (before PowerPointApi 1.8)', async () => {
      const d = layoutDeck();
      installed = install(d, { requirements: { PowerPointApi: 1.5 } });
      await new PowerPointBridge().actuate(insertSlide({ slide: { title: 'T', bullets: ['B'] } }));
      expect(d.addOptions).toEqual([undefined]);
    });
  });

  it('includes the PowerPoint error code when a write is uncertain', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [
      { text: '', type: 'Placeholder', placeholderType: 'Title', refuseText: true },
    ];
    installed = install(d);
    const sync = vi.spyOn(FakeContext.prototype, 'sync').mockImplementation(function (
      this: FakeContext,
    ) {
      this.syncs += 1;
      const error = pendingHostError;
      pendingHostError = undefined;
      return error
        ? Promise.reject(Object.assign(new Error(error), { code: 'GeneralException' }))
        : Promise.resolve();
    });
    try {
      const res = await new PowerPointBridge().actuate(
        insertSlide({ slide: { title: 'T', bullets: [] } }),
      );
      expect(res.error?.code).toBe('outcome_unknown');
      expect(res.error?.message).toContain('(PowerPoint: GeneralException)');
    } finally {
      sync.mockRestore();
    }
  });

  it('never writes onto another slide when the new one cannot be identified (co-author race)', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    installed = install(d);
    // A co-author's slide lands in the same moment: two new ids appear, so neither is trusted.
    const add = vi.spyOn(FakeSlideCollection.prototype, 'add').mockImplementation(() => {
      d.slides.push({ id: 'coauthor-slide', shapes: [{ text: 'Theirs' }] });
      d.slides.push({ id: 'ours', shapes: [] });
    });
    try {
      const res = await new PowerPointBridge().actuate(
        insertSlide({ slide: { title: 'T', bullets: ['B'] } }),
      );
      expect(res).toMatchObject({ ok: false, error: { code: 'outcome_unknown' } });
      expect(d.slides.find((slide) => slide.id === 'coauthor-slide')?.shapes[0]?.text).toBe(
        'Theirs',
      );
    } finally {
      add.mockRestore();
    }
  });

  it('fails closed, writing nothing, when a layout slide never finishes (PowerPoint web)', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout'] });
    try {
      const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
      d.layouts = [
        {
          id: 'L-content',
          type: 'Object',
          shapes: [
            { text: '', type: 'Placeholder', placeholderType: 'Title' },
            { text: '', type: 'Placeholder', placeholderType: 'Body' },
          ],
        },
      ];
      d.settleIdsAfterSyncs = 10_000;
      d.layoutShapesOnSettle = true;
      installed = install(d);
      const pending = new PowerPointBridge().actuate(
        insertSlide({ slide: { title: 'Monitoring', bullets: ['Drift'] } }),
      );
      await vi.runAllTimersAsync();
      const res = await pending;
      // The empty slide is removed again, so trying again cannot leave a duplicate.
      expect(res.ok).toBe(false);
      expect(res.error?.code).toBe('slide_not_ready');
      expect(res.error?.message).toContain('nothing was added');
      expect(res.recoveryPending).toBeUndefined();
      expect(d.slides.map((slide) => slide.id)).toEqual(['s1']);
      expect(d.addedShapes).toEqual([]);
    } finally {
      vi.useRealTimers();
    }
  });

  it('fails closed when a co-author shifts the deck before the new slide is written', async () => {
    // s1 has two shapes, so a write that shifted onto it would land silently — the dangerous case.
    const d = deck([{ id: 's1', shapes: [{ text: 'Mine' }, { text: 'Also mine' }] }], [0]);
    installed = install(d);
    // The placeholder-type read runs after the identity check and before the text writes: a
    // co-author's slide lands at the top right then, so position 1 is no longer our slide.
    const real = Object.getOwnPropertyDescriptor(FakeShape.prototype, 'placeholderFormat')!;
    let shifted = false;
    const spy = vi
      .spyOn(FakeShape.prototype, 'placeholderFormat', 'get')
      .mockImplementation(function (this: FakeShape) {
        if (!shifted && d.slides.length === 2) {
          shifted = true;
          d.slides.unshift({ id: 'coauthor-slide', shapes: [{ text: 'Theirs' }] });
        }
        return real.get!.call(this);
      });
    try {
      const res = await new PowerPointBridge().actuate(
        insertSlide({ slide: { title: 'Ours', bullets: ['B'] } }),
      );
      expect(shifted).toBe(true);
      expect(res.ok).toBe(false);
      expect(res.error?.code).toBe('outcome_unknown');
      expect(res.error?.message).toContain('may have landed on another slide');
    } finally {
      spy.mockRestore();
    }
  });

  it('fails closed when the new slide changes id a second time while it settles', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout'] });
    try {
      const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
      // A layout whose shapes never show, so the bridge keeps probing after the id settles.
      d.layouts = [
        {
          id: 'L-content',
          type: 'Object',
          shapes: [
            { text: '', type: 'Placeholder', placeholderType: 'Title' },
            { text: '', type: 'Placeholder', placeholderType: 'Body' },
          ],
        },
      ];
      d.settleIdsAfterSyncs = 3;
      d.layoutShapesOnSettle = true;
      // After the one provisional → real swap, our slide is deleted and another new one takes its
      // place: same position, same count, one new id — but a second id change.
      d.onSync = () => {
        const slide = d.slides[1];
        // Its shapes never show, so the bridge keeps probing after the id settles.
        if (slide?.pendingShapes) slide.pendingShapes = [];
        if (slide?.id === 'sim-slide-2') slide.id = 'someone-elses-slide';
      };
      installed = install(d);
      const pending = new PowerPointBridge().actuate(
        insertSlide({ slide: { title: 'Ours', bullets: ['B'] } }),
      );
      await vi.runAllTimersAsync();
      const res = await pending;
      expect(res.ok).toBe(false);
      expect(res.error?.message).toContain('The deck changed while the slide was being added');
    } finally {
      vi.useRealTimers();
    }
  });

  it('reports an unconfirmed write as outcome-unknown, never as success', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [
      { text: '', type: 'Placeholder', placeholderType: 'Title', refuseText: true },
    ];
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'T', bullets: [] } }),
    );
    expect(res.ok).toBe(false);
    expect(res.error?.code).toBe('outcome_unknown');
  });
});

describe('PowerPointBridge shape and slide references (tests 48–51)', () => {
  it('changes only the title when a shape command names slide 1 / shape title', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', 'title', 'FY26 Plan'));
    expect(res).toMatchObject({ ok: true, location: 'shape:s1:s1-shape-0' });
    expect(d.slides[0]?.shapes.map((shape) => shape.text)).toEqual([
      'FY26 Plan',
      '99.5% contracted\nMonthly window',
    ]);
  });

  it('formats the title shape resolved by slide number and "title"', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: '1', shapeId: 'title' },
        shapeFormat: { fill: '#0000FF', font: { color: '#FFFFFF' } },
      }),
    );
    expect(res.ok).toBe(true);
    expect(d.slides[0]?.shapes[0]).toMatchObject({
      fillColor: '#0000FF',
      font: { color: '#FFFFFF' },
    });
    expect(d.slides[0]?.shapes[1]?.fillColor).toBeUndefined();
  });

  it('names the valid shapes when a shape reference does not resolve, and writes nothing', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', 'nope', 'x'));
    expect(res).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('pp:shape:1:title');
    expect(res.error?.message).toContain('id s1-shape-0 (TextBox, title: "SLA Terms")');
    expect(res.error?.message).toContain(
      'id s1-shape-1 (TextBox, "99.5% contracted Monthly window")',
    );
    expect(d.slides[0]?.shapes[0]?.text).toBe('SLA Terms');
  });

  it('marks shapes without text in the not-found listing', async () => {
    installed = install(
      deck(
        [
          {
            id: 's1',
            shapes: [
              { text: '', type: 'GeometricShape', zOrderCalls: [] },
              { text: 'What is Tokenomics?', zOrderCalls: [] },
            ],
          },
        ],
        [0],
      ),
    );
    const res = await new PowerPointBridge().actuate(setShapeText('1', '1', 'Q4 plan'));
    expect(res.error?.message).toContain(
      'id s1-shape-0 (GeometricShape, no text), id s1-shape-1 (TextBox, title: "What is Tokenomics?")',
    );
  });

  it('takes the largest text as the title when the slide has no title placeholder', async () => {
    // Live (a Google Slides export): a 12 pt "TOKENOMICS / 01" label sits above the 28.5 pt title,
    // and "first shape with text" retitled the label.
    const d = deck(
      [
        {
          id: 's2',
          shapes: [
            { text: 'TOKENOMICS / 01', top: 36, firstCharSize: 12, zOrderCalls: [] },
            { text: '', type: 'GeometricShape', top: 57, zOrderCalls: [] },
            {
              text: 'What do we mean by Tokenomics?',
              top: 76,
              firstCharSize: 28.5,
              zOrderCalls: [],
            },
            { text: 'NOT JUST A FORMULA', top: 153, firstCharSize: 8.25, zOrderCalls: [] },
          ],
        },
      ],
      [0],
    );
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', 'title', 'Q4 plan'));
    expect(res).toMatchObject({ ok: true, location: 'shape:s2:s2-shape-2' });
    expect(d.slides[0]?.shapes.map((shape) => shape.text)).toEqual([
      'TOKENOMICS / 01',
      '',
      'Q4 plan',
      'NOT JUST A FORMULA',
    ]);
  });

  it('breaks a title font-size tie by the topmost shape', async () => {
    const d = deck(
      [
        {
          id: 's1',
          shapes: [
            { text: 'Lower', top: 200, firstCharSize: 30, zOrderCalls: [] },
            { text: 'Upper', top: 40, firstCharSize: 30, zOrderCalls: [] },
          ],
        },
      ],
      [0],
    );
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', 'title', 'New'));
    expect(res).toMatchObject({ ok: true, location: 'shape:s1:s1-shape-1' });
  });

  it('rejects a format-shape with no formatting instead of reporting a no-op as applied', async () => {
    // Live: `/format-shape shape:257#0:6 text="…"` came back "applied" having changed nothing.
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      formatShape({ target: { slideId: '1', shapeId: 'title' }, shapeFormat: {} }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'no_format' } });
    expect(res.error?.message).toContain('shape pp:shape:<slide>:<shape> "text"');
  });

  it('does not treat a shape number as a position (ids are small numbers too)', async () => {
    // Live: `shape pp:shape:1:1 "FY26 Plan"` resolved "1" as the 1st shape — the logo image.
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', '1', 'FY26 Plan'));
    expect(res).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('pp:shape:1:title');
    expect(d.slides[0]?.shapes.map((shape) => shape.text)).toEqual([
      'SLA Terms',
      '99.5% contracted\nMonthly window',
    ]);
  });

  it('refuses to write text into a picture, with a message naming shape=title', async () => {
    const d = deck(
      [{ id: 's1', shapes: [{ text: '', type: 'Image' }, { text: 'Deck title' }] }],
      [0],
    );
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', 's1-shape-0', 'x'));
    expect(res).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('is a Image and has no text');
    expect(d.slides[0]?.shapes[1]?.text).toBe('Deck title');
  });

  it('colours the header row with headerFill and white bold text (UAT #16)', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('table-header-fill'),
      kind: 'add-table-slide',
      surface: 'powerpoint',
      params: {
        target: { slideId: '2' },
        tableGrid: {
          hasHeaders: true,
          headerFill: '#C8102E',
          rows: [
            ['No', 'Dimensi'],
            ['1', 'Regulasi'],
          ],
        },
      },
    });
    expect(res.ok).toBe(true);
    const options = d.addedShapes.at(-1)?.options as PowerPoint.TableAddOptions;
    expect(options.specificCellProperties).toEqual([
      [
        { fill: { color: '#C8102E' }, font: { color: '#FFFFFF', bold: true, size: 14 } },
        { fill: { color: '#C8102E' }, font: { color: '#FFFFFF', bold: true, size: 14 } },
      ],
      [{}, {}],
    ]);
  });

  it('puts a table below the bullets of a full slide, inside the slide (real PowerPoint, 2026-10-17)', async () => {
    const d = deck(
      [
        {
          id: 's4',
          shapes: [
            {
              text: 'Analisa',
              type: 'Placeholder',
              placeholderType: 'Title',
              left: 48,
              top: 20,
              width: 624,
              height: 60,
            },
            {
              text: 'Bank perlu…\nMendapat pengakuan…\nPemilihan Cloud Provider…',
              type: 'Placeholder',
              placeholderType: 'Body',
              left: 48,
              top: 96,
              width: 624,
              height: 280,
            },
          ],
        },
      ],
      [0],
    );
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addTableSlide({
        target: { slideId: 's4' },
        // The position the model copied from the example: on top of the bullets.
        tableGrid: {
          hasHeaders: true,
          rows: [
            ['No', 'Dimensi', 'Deskripsi'],
            ['1', 'Keamanan', 'Keamanan data tingkat tinggi'],
            ['2', 'Integrasi', 'Integrasi mudah dengan sistem internal'],
            ['3', 'Kepatuhan', 'Sesuai regulasi OJK'],
          ],
          left: 72,
          top: 110,
          width: 560,
          height: 200,
        },
      }),
    );
    expect(res.ok).toBe(true);
    const options = d.addedShapes.at(-1)?.options as PowerPoint.TableAddOptions;
    const body = d.slides[0]!.shapes[1]!;
    // Below the (shortened) bullets, never on them; the whole table inside the 405 pt slide.
    expect(options.top!).toBeGreaterThanOrEqual(body.top! + body.height!);
    expect(body.autoSize).toBe('AutoSizeTextToFitShape');
    const widths = options.columns!.map((c) => c.columnWidth!);
    expect(widths.reduce((a, b) => a + b, 0)).toBeCloseTo(options.width!, 3);
    expect(widths[0]).toBeLessThan(widths[2]!); // a narrow "No" column
    const rows = [
      ['No', 'Dimensi', 'Deskripsi'],
      ['1', 'Keamanan', 'Keamanan data tingkat tinggi'],
      ['2', 'Integrasi', 'Integrasi mudah dengan sistem internal'],
      ['3', 'Kepatuhan', 'Sesuai regulasi OJK'],
    ];
    expect(options.top! + tableHeight(rows, widths)).toBeLessThanOrEqual(405);
  });

  it('rejects title= on an existing slide instead of silently ignoring it', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('table-title-existing'),
      kind: 'add-table-slide',
      surface: 'powerpoint',
      params: {
        target: { slideId: '2' },
        slide: { title: 'Ignored?', bullets: [] },
        tableGrid: { hasHeaders: false, rows: [['a']] },
      },
    });
    expect(res).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('only used with slide=new');
    expect(d.addedShapes).toEqual([]);
  });

  it('rejects a text box with no text, and a shape with no position, before writing', async () => {
    // Live: `/add-shape slide:2 {"text": …, "position": "bottom"}` → an empty box at the top-left.
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const bridge = new PowerPointBridge();
    const empty = await bridge.actuate(
      addShape({
        target: { slideId: '2' },
        shape: { shapeType: 'textBox', left: 72, top: 300, width: 300, height: 40 },
      }),
    );
    expect(empty).toMatchObject({ ok: false, error: { code: 'no_text' } });
    const unplaced = await bridge.actuate(
      addShape({ target: { slideId: '2' }, shape: { shapeType: 'textBox', text: 'Draft' } }),
    );
    expect(unplaced).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(unplaced.error?.message).toContain('missing left, top, width, height');
    expect(unplaced.error?.message).toContain('The slide is 720 × 405 pt');
    expect(d.addedShapes).toEqual([]);
  });

  it('accepts pp:slide:<id> and rejects an off-slide position with the slide size', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const bridge = new PowerPointBridge();
    const ok = await bridge.actuate(
      addShape({
        target: { slideId: 'pp:slide:s2' },
        shape: { shapeType: 'textBox', text: 'Draft', left: 72, top: 330, width: 400, height: 40 },
      }),
    );
    expect(ok.ok).toBe(true);
    expect(d.addedShapes[0]?.slideId).toBe('s2');
    const off = await bridge.actuate(
      addShape({
        target: { slideId: '2' },
        shape: { shapeType: 'textBox', text: 'Draft', left: 100, top: 500, width: 300, height: 50 },
      }),
    );
    expect(off).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(off.error?.message).toContain('outside the slide, which is 720 × 405 pt');
    expect(d.addedShapes).toHaveLength(1);
  });

  it('places a slide=new table slide at the requested position (at=2)', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const before = d.slides.map((s) => s.id);
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('table-at'),
      kind: 'add-table-slide',
      surface: 'powerpoint',
      params: {
        target: { slideId: 'new', slideIndex: 1 },
        slide: { title: 'Key risks', bullets: [] },
        tableGrid: { hasHeaders: true, rows: [['Risk', 'Owner']] },
      },
    });
    expect(res).toMatchObject({ ok: true });
    expect(res.error).toBeUndefined();
    expect(d.slides[1]?.shapes[0]?.text).toBe('Key risks');
    expect(d.slides[0]?.id).toBe(before[0]);
    // Undo still removes the created slide, now at position 2.
    expect(res.inverse).toMatchObject({ op: 'delete-object', objectType: 'slide' });
  });

  it('fails closed when the new slide changes id again after its one swap (security review)', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout'] });
    try {
      const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
      d.layouts = [
        {
          id: 'L-title',
          type: 'TitleOnly',
          shapes: [{ text: '', type: 'Placeholder', placeholderType: 'Title' }],
        },
      ];
      d.settleIdsAfterSyncs = 3;
      d.layoutShapesOnSettle = true;
      // After the table lands, our slide is replaced by a co-author's at the same position.
      d.onSync = () => {
        const slide = d.slides[1];
        if (d.addedShapes.some((s) => s.kind === 'table') && slide?.id === 'sim-slide-2')
          slide.id = 'someone-elses-slide';
      };
      installed = install(d);
      const pending = new PowerPointBridge().actuate({
        changeId: asChangeId('table-swap'),
        kind: 'add-table-slide',
        surface: 'powerpoint',
        params: {
          target: { slideId: 'new' },
          slide: { title: 'Key risks', bullets: [] },
          tableGrid: { hasHeaders: true, rows: [['Risk', 'Owner']] },
        },
      });
      await vi.runAllTimersAsync();
      const res = await pending;
      // Before the fix: ok, with an undo that deletes 'someone-elses-slide'.
      expect(res.ok).toBe(false);
      expect(res.error?.code).toBe('outcome_unknown');
      expect(res.inverse).toBeUndefined();
    } finally {
      vi.useRealTimers();
    }
  });

  it('refuses at= on an existing slide before writing anything', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('table-at-existing'),
      kind: 'add-table-slide',
      surface: 'powerpoint',
      params: {
        target: { slideId: '2', slideIndex: 0 },
        tableGrid: { hasHeaders: true, rows: [['Risk', 'Owner']] },
      },
    });
    expect(res.ok).toBe(false);
    expect(res.error?.message).toContain('at= is only used with slide=new');
    expect(d.addedShapes).toHaveLength(0);
  });

  it('creates a titled slide for /add-table-slide slide=new and puts the table on it', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('table-new'),
      kind: 'add-table-slide',
      surface: 'powerpoint',
      params: {
        target: { slideId: 'new' },
        slide: { title: 'Key risks', bullets: [] },
        tableGrid: {
          hasHeaders: true,
          rows: [
            ['Risk', 'Owner'],
            ['Vendor delay', 'Pat'],
          ],
        },
      },
    });
    expect(res).toMatchObject({ ok: true });
    expect(res.location).toMatch(/^shape:sim-slide-4:/);
    // With no position given, the table goes below the title, inside the 720 × 405 slide.
    const tableAdd = d.addedShapes.find((added) => added.kind === 'table');
    expect(tableAdd?.options).toMatchObject({ left: 48, top: 109, width: 624 });
    // Undo removes the slide this command created (and the table with it).
    expect(res.inverse).toEqual({ op: 'delete-object', objectType: 'slide', name: 'sim-slide-4' });
    const created = d.slides[3];
    expect(created?.shapes[0]?.text).toBe('Key risks');
    expect(created?.shapes.find((shape) => shape.tableGrid)?.tableGrid).toEqual([
      ['Risk', 'Owner'],
      ['Vendor delay', 'Pat'],
    ]);
    // The existing last slide was not touched.
    expect(d.slides[2]?.shapes.map((shape) => shape.text)).toEqual([
      'Risk Summary',
      'SLA gap flagged',
    ]);
  });

  it('adds a title text box for slide=new on a layout without placeholders', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    d.newSlideShapes = [];
    installed = install(d);
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('table-new-blank'),
      kind: 'add-table-slide',
      surface: 'powerpoint',
      params: {
        target: { slideId: 'new' },
        slide: { title: 'Key risks', bullets: [] },
        tableGrid: { hasHeaders: false, rows: [['Vendor delay', 'Pat']] },
      },
    });
    expect(res.ok).toBe(true);
    expect(d.slides[3]?.shapes.map((shape) => shape.type ?? 'TextBox')).toEqual([
      'TextBox',
      'Table',
    ]);
    expect(d.slides[3]?.shapes[0]?.text).toBe('Key risks');
  });
});

describe('PowerPointBridge.actuate set-shape-text', () => {
  it.each([
    ['s1', 'no-such-shape', 'shape_not_found'],
    ['no-such-slide', 's1-shape-0', 'slide_not_found'],
  ])('tags a %s:%s conflict with a content-free reason', async (slideId, shapeId, reason) => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const res = await new PowerPointBridge().actuate(setShapeText(slideId, shapeId, 'New'));
    expect(res.error).toMatchObject({ code: 'target_conflict', reason });
  });

  // Syncs: 1 resolve slide, 2 resolve shape, 3 shape exists?, 4 prior text read, 5 the write.
  it.each([1, 2, 3, 4, 5])(
    'distinguishes a failed pre-read from a dispatched text write (sync %s)',
    async (failAt) => {
      const d = deck(SAMPLE_SLIDES, [0]);
      installed = install(d);
      const original = FakeContext.prototype.sync;
      let calls = 0;
      const sync = vi.spyOn(FakeContext.prototype, 'sync').mockImplementation(function (
        this: FakeContext,
      ) {
        if (++calls === failAt) return Promise.reject(new Error('Office disconnected'));
        return original.call(this);
      });
      try {
        const result = await new PowerPointBridge().actuate(
          setShapeText('s2', 's2-shape-1', 'New text'),
        );
        const beforeWrite = failAt < 5;
        expect(result).toMatchObject({
          ok: false,
          error: { code: beforeWrite ? 'target_conflict' : 'outcome_unknown' },
        });
        expect(result.recoveryPending).toBe(beforeWrite ? undefined : true);
        expect(d.slides[1]?.shapes[1]?.text).toBe(beforeWrite ? 'Pat, Sam' : 'New text');
      } finally {
        sync.mockRestore();
      }
    },
  );

  it('turns each line of shape text into its own paragraph (literal \\n or a real newline)', async () => {
    // A \n is a line break inside one paragraph: a typed list showed as one run-on paragraph.
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    await new PowerPointBridge().actuate(setShapeText('s2', 's2-shape-1', 'Pat\\nSam\nLee'));
    expect(d.slides[1]?.shapes[1]?.text).toBe('Pat\rSam\rLee');
  });

  it('writes a numbered list into a shape without the placeholder bullet (UAT daftar isi)', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    await new PowerPointBridge().actuate(
      setShapeText(
        's2',
        's2-shape-1',
        '1. Latar belakang\\n2. Governance principles\\n3. Monitoring',
      ),
    );
    const shape = d.slides[1]?.shapes[1];
    expect(shape?.text).toBe('1. Latar belakang\r2. Governance principles\r3. Monitoring');
    expect(shape?.bulletsVisible).toBe(false);
  });

  it('keeps the bullets for plain lines', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    await new PowerPointBridge().actuate(setShapeText('s2', 's2-shape-1', 'Pat\\nSam'));
    expect(d.slides[1]?.shapes[1]?.bulletsVisible).toBeUndefined();
  });

  it('rewrites one addressed shape and records the inverse text', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      setShapeText('s2', 's2-shape-1', 'Pat, Sam, and Lee'),
    );

    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('shape-1'),
      kind: 'set-shape-text',
      location: 'shape:s2:s2-shape-1',
      inverse: {
        op: 'restore-text',
        anchor: 'pp:shape:s2:s2-shape-1',
        priorText: 'Pat, Sam',
      },
    });
    expect(d.slides[1]?.shapes[1]?.text).toBe('Pat, Sam, and Lee');
  });

  it('fails closed when the addressed shape no longer exists', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      setShapeText('s2', 's2-shape-99', 'Should not apply'),
    );

    expect(res).toMatchObject({
      ok: false,
      kind: 'set-shape-text',
      degraded: true,
      error: { code: 'target_conflict' },
    });
    expect(d.slides[1]?.shapes[1]?.text).toBe('Pat, Sam');
  });

  it('fails before touching the host when the required shape text API is unavailable', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.3 } });
    const res = await new PowerPointBridge().actuate(
      setShapeText('s2', 's2-shape-1', 'Unsupported'),
    );

    expect(res).toMatchObject({
      ok: false,
      kind: 'set-shape-text',
      error: { code: 'unsupported' },
    });
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

/* ───────────────────────────── actuate: add-shape ───────────────────────────── */

function addShape(params: ActuationRequest['params'], id = 'add-shape-1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'add-shape', surface: 'powerpoint', params };
}

describe('PowerPointBridge.actuate add-shape', () => {
  describe('text box over empty layout placeholders (UAT 2026-10-16 daftar isi)', () => {
    const blankContentSlide = (title = '', body = ''): DeckSeed =>
      deck(
        [
          { id: 's1', shapes: [{ text: 'Agenda' }] },
          {
            id: 's2',
            shapes: [
              { text: title, type: 'Placeholder', placeholderType: 'Title' },
              { text: body, type: 'Placeholder', placeholderType: 'Body' },
            ],
          },
        ],
        [1],
      );
    const toc = addShape({
      target: { slideId: '2' },
      shape: {
        shapeType: 'textBox',
        text: 'Table of Contents\\n1. Background',
        left: 72,
        top: 120,
        width: 560,
        height: 250,
      },
    });

    it('refuses before writing and names the placeholders to fill', async () => {
      const d = blankContentSlide();
      installed = install(d);
      const res = await new PowerPointBridge().actuate(toc);
      expect(res.ok).toBe(false);
      expect(res.error?.code).toBe('use_placeholders');
      expect(res.error?.message).toContain('shape pp:shape:2:title');
      expect(res.error?.message).toContain('shape pp:shape:2:s2-shape-1');
      expect(d.slides[1]?.shapes).toHaveLength(2);
    });

    it('adds the text box once the placeholders hold text', async () => {
      const d = blankContentSlide('Daftar Isi', '1. Latar belakang');
      installed = install(d);
      const res = await new PowerPointBridge().actuate(toc);
      expect(res.ok).toBe(true);
      expect(d.slides[1]?.shapes).toHaveLength(3);
    });

    it('adds the text box on a slide without placeholders', async () => {
      const d = deck(
        [
          { id: 's1', shapes: [] },
          { id: 's2', shapes: [{ text: 'Note' }] },
        ],
        [1],
      );
      installed = install(d);
      const res = await new PowerPointBridge().actuate(toc);
      expect(res.ok).toBe(true);
    });
  });

  it('reports unknown outcome when shape creation landed but its sync failed', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    // Every read-only sync succeeds; the first sync after the text box is queued fails.
    let queued = false;
    const realAdd = FakeShapeCollection.prototype.addTextBox;
    const add = vi.spyOn(FakeShapeCollection.prototype, 'addTextBox').mockImplementation(function (
      this: FakeShapeCollection,
      ...args
    ) {
      queued = true;
      return realAdd.apply(this, args);
    });
    const realSync = FakeContext.prototype.sync;
    const sync = vi.spyOn(FakeContext.prototype, 'sync').mockImplementation(function (
      this: FakeContext,
    ) {
      return queued ? Promise.reject(new Error('Office disconnected')) : realSync.call(this);
    });
    try {
      const result = await new PowerPointBridge().actuate(
        addShape({
          target: { slideId: 's1' },
          shape: {
            shapeType: 'textBox',
            text: 'Created',
            left: 72,
            top: 300,
            width: 300,
            height: 40,
          },
        }),
      );
      expect(result).toMatchObject({
        ok: false,
        recoveryPending: true,
        error: { code: 'outcome_unknown' },
      });
      expect(d.slides[0]?.shapes).toHaveLength(3);
    } finally {
      sync.mockRestore();
      add.mockRestore();
    }
  });

  it('adds a text box to the addressed slide, recording the minted shape id as delete-object inverse', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape(
        {
          target: { slideId: 's2' },
          shape: {
            shapeType: 'textBox',
            text: 'Key risk',
            left: 72,
            top: 96,
            width: 280,
            height: 80,
          },
        },
        'chg-add-textbox',
      ),
    );

    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-add-textbox'),
      kind: 'add-shape',
      location: 'shape:s2:s2-shape-2',
      inverse: { op: 'delete-object', objectType: 'shape', name: 's2-shape-2' },
    });
    // The text box landed on slide s2 with its text and explicit geometry.
    expect(d.addedShapes).toEqual([
      {
        slideId: 's2',
        kind: 'textBox',
        detail: 'Key risk',
        options: { left: 72, top: 96, width: 280, height: 80 },
      },
    ]);
    expect(d.slides[1]?.shapes).toHaveLength(3);
    expect(d.slides[1]?.shapes[2]?.text).toBe('Key risk');
  });

  it('adds a filled geometric shape with a whitelisted geometry literal', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape({
        target: { slideId: 's1' },
        shape: {
          shapeType: 'geometric',
          geometryType: 'Rectangle',
          fill: '#0F6CBD',
          left: 72,
          top: 96,
          width: 120,
          height: 60,
        },
      }),
    );

    expect(res.ok).toBe(true);
    expect(res.location).toBe('shape:s1:s1-shape-2');
    expect(d.addedShapes[0]).toMatchObject({ kind: 'geometric', detail: 'Rectangle' });
    expect(d.slides[0]?.shapes[2]?.fillColor).toBe('#0F6CBD');
  });

  it('writes text= inside a geometric shape instead of dropping it', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape({
        target: { slideId: 's2' },
        shape: {
          shapeType: 'geometric',
          geometryType: 'triangle',
          text: 'AI CoE',
          fill: '#0000FF',
          left: 560,
          top: 20,
          width: 140,
          height: 60,
        },
      }),
    );

    expect(res.ok).toBe(true);
    expect(d.addedShapes[0]).toMatchObject({ kind: 'geometric', detail: 'Triangle' });
    expect(d.slides[1]?.shapes[2]).toMatchObject({ text: 'AI CoE', fillColor: '#0000FF' });
  });

  it('maps the contract connector name onto the host ConnectorType for lines', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape({
        target: { slideId: 's3' },
        shape: {
          shapeType: 'line',
          connectorType: 'elbow',
          left: 72,
          top: 300,
          width: 300,
          height: 40,
        },
      }),
    );

    expect(res.ok).toBe(true);
    expect(d.addedShapes).toHaveLength(1);
    expect(d.addedShapes[0]).toMatchObject({ kind: 'line', detail: 'Elbow' });
  });

  it('fails closed without touching the host when the target or shape params are missing', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const bridge = new PowerPointBridge();

    const noSlide = await bridge.actuate(addShape({ shape: { shapeType: 'textBox', text: 'x' } }));
    expect(noSlide).toMatchObject({ ok: false, error: { code: 'no_target' } });

    const noShape = await bridge.actuate(addShape({ target: { slideId: 's2' } }));
    expect(noShape).toMatchObject({ ok: false, error: { code: 'no_shape' } });

    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('rejects a non-whitelisted geometric geometry as unsupported before touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const res = await new PowerPointBridge().actuate(
      addShape({
        target: { slideId: 's2' },
        shape: { shapeType: 'geometric', geometryType: 'PortalGun' },
      }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it.each([
    ['last', 's3'],
    ['LAST', 's3'],
    ['2', 's2'],
    ['slide:s1', 's1'],
  ])('resolves slide=%s to the right slide before writing', async (ref, slideId) => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape({
        target: { slideId: ref },
        shape: { shapeType: 'textBox', text: 'x', left: 72, top: 300, width: 300, height: 40 },
      }),
    );
    expect(res.ok).toBe(true);
    expect(d.addedShapes[0]?.slideId).toBe(slideId);
    expect(res.location).toMatch(new RegExp(`^shape:${slideId}:`));
  });

  it('points a missing-slide add-shape / add-table-slide at the slide command', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const bridge = new PowerPointBridge();
    const table = await bridge.actuate(
      addTableSlide({ tableGrid: { hasHeaders: false, rows: [['hire 5 engineers']] } }),
    );
    const shape = await bridge.actuate(addShape({ shape: { shapeType: 'textBox', text: 'x' } }));
    for (const res of [table, shape]) {
      expect(res).toMatchObject({ ok: false, error: { code: 'no_target' } });
      expect(res.error?.message).toContain('slide "Title" "bullet"');
      expect(res.error?.message).toContain('a slide id from the outline, a slide number, or last');
      // The runtime escapes < > in results; keep the hint free of them so it reads cleanly.
      expect(res.error?.message).not.toMatch(/[<>]/);
    }
    expect(table.error?.message).toContain('slide=new title=');
  });

  it('fails with the valid slides, not "outcome uncertain", when the slide does not exist', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape({ target: { slideId: 'nope' }, shape: { shapeType: 'textBox', text: 'x' } }),
    );
    expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('Slide "nope" was not found');
    expect(res.error?.message).toContain('or "last"');
    expect(res.error?.message).toContain('s1, s2, s3');
    expect(d.addedShapes).toEqual([]);
  });

  it('degrades to unsupported on a pre-1.4 host without touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.3 } });
    const res = await new PowerPointBridge().actuate(
      addShape({ target: { slideId: 's2' }, shape: { shapeType: 'textBox', text: 'x' } }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.error?.message).toContain('PowerPointApi 1.4');
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

/* ───────────────────────────── actuate: format-shape ───────────────────────────── */

function formatShape(params: ActuationRequest['params'], id = 'format-shape-1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'format-shape', surface: 'powerpoint', params };
}

function styledDeck(): DeckSeed {
  return deck([
    {
      id: 's2',
      shapes: [
        { text: 'Roster', zOrderCalls: [] },
        {
          text: 'Pat, Sam',
          zOrderCalls: [],
          fillColor: '#FFFFFF',
          lineColor: '#000000',
          font: {
            bold: false,
            italic: false,
            underline: 'None',
            color: '#111111',
            size: 18,
            name: 'Arial',
          },
        },
      ],
    },
  ]);
}

describe('PowerPointBridge.actuate format-shape', () => {
  it('lists shape ids before the title hint when a shape name is not found', async () => {
    // Live 2026-10-09: told "For the slide title use title" first, the model recoloured the title.
    installed = install(styledDeck());
    const res = await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: '1', shapeId: 'myRectangle' },
        shapeFormat: { fill: '#0000FF' },
      }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    const message = res.error?.message ?? '';
    expect(message.indexOf('Use one of its shape ids')).toBeLessThan(message.indexOf('title'));
    expect(message).toContain('the id in its add-shape result');
  });

  it('retains uncertainty when one format facet landed before the next prior-state read failed', async () => {
    const d = styledDeck();
    installed = install(d);
    const original = FakeContext.prototype.sync;
    let calls = 0;
    const sync = vi.spyOn(FakeContext.prototype, 'sync').mockImplementation(function (
      this: FakeContext,
    ) {
      // Syncs 1–2 resolve the slide and shape; 3 exists?, 4 fill prior read, 5 line prior read
      // (which also carries the queued fill write).
      if (++calls === 5) return Promise.reject(new Error('Office disconnected'));
      return original.call(this);
    });
    try {
      const result = await new PowerPointBridge().actuate(
        formatShape({
          target: { slideId: 's2', shapeId: 's2-shape-1' },
          shapeFormat: { fill: '#123456', line: '#654321' },
        }),
      );
      expect(result).toMatchObject({
        ok: false,
        recoveryPending: true,
        error: { code: 'outcome_unknown' },
      });
      expect(d.slides[0]?.shapes[1]?.fillColor).toBe('#123456');
      expect(d.slides[0]?.shapes[1]?.lineColor).toBe('#000000');
    } finally {
      sync.mockRestore();
    }
  });

  it('applies fill/line/font deltas and records only the touched prior fields as inverse', async () => {
    const d = styledDeck();
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      formatShape(
        {
          target: { slideId: 's2', shapeId: 's2-shape-1' },
          shapeFormat: {
            fill: '#0F6CBD',
            line: '#FF0000',
            font: { bold: true, italic: true, color: '#222222', size: 24, name: 'Verdana' },
          },
        },
        'chg-format',
      ),
    );

    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-format'),
      kind: 'format-shape',
      location: 'shape:s2:s2-shape-1',
      inverse: {
        op: 'restore-shape-format',
        shapeId: 's2-shape-1',
        prior: {
          fill: '#FFFFFF',
          line: '#000000',
          'font.bold': 'false',
          'font.italic': 'false',
          'font.color': '#111111',
          'font.size': '18',
          'font.name': 'Arial',
        },
      },
    });
    const shape = d.slides[0]?.shapes[1];
    expect(shape?.fillColor).toBe('#0F6CBD');
    expect(shape?.lineColor).toBe('#FF0000');
    expect(shape?.font).toMatchObject({
      bold: true,
      italic: true,
      color: '#222222',
      size: 24,
      name: 'Verdana',
      underline: 'None', // untouched
    });
  });

  it('maps the boolean underline onto the host Single/None style literals', async () => {
    const d = styledDeck();
    installed = install(d);
    await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-1' },
        shapeFormat: { font: { underline: true } },
      }),
    );
    expect(d.slides[0]?.shapes[1]?.font?.underline).toBe('Single');
  });

  it('moves the shape in z-order on a 1.8 host', async () => {
    const d = styledDeck();
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-1' },
        shapeFormat: { zOrder: 'front' },
      }),
    );

    expect(res.ok).toBe(true);
    expect(res.inverse).toMatchObject({ op: 'restore-shape-format', prior: {} });
    expect(d.slides[0]?.shapes[1]?.zOrderCalls).toEqual(['BringToFront']);
  });

  it('rejects z-order on a pre-1.8 host without touching the host', async () => {
    installed = install(styledDeck(), { requirements: { PowerPointApi: 1.7 } });
    const res = await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-1' },
        shapeFormat: { zOrder: 'back' },
      }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.error?.message).toContain('PowerPointApi 1.8');
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('fails closed when the addressed shape or formatting params are missing', async () => {
    const d = styledDeck();
    installed = install(d);
    const bridge = new PowerPointBridge();

    const noTarget = await bridge.actuate(formatShape({ shapeFormat: { fill: '#000000' } }));
    expect(noTarget).toMatchObject({ ok: false, error: { code: 'no_target' } });

    const noFormat = await bridge.actuate(
      formatShape({ target: { slideId: 's2', shapeId: 's2-shape-1' } }),
    );
    expect(noFormat).toMatchObject({ ok: false, error: { code: 'no_format' } });

    // Both param failures rejected the request before any host object was touched.
    expect(installed.ctxRef.ctx).toBeUndefined();

    // A stale shape id degrades instead of formatting some other shape.
    const stale = await bridge.actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-99' },
        shapeFormat: { fill: '#000000' },
      }),
    );
    expect(stale).toMatchObject({
      ok: false,
      degraded: true,
      error: { code: 'target_conflict' },
    });
    expect(d.slides[0]?.shapes[1]?.fillColor).toBe('#FFFFFF');
  });

  it('degrades to unsupported on a pre-1.4 host without touching the host', async () => {
    installed = install(styledDeck(), { requirements: { PowerPointApi: 1.3 } });
    const res = await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-1' },
        shapeFormat: { fill: '#000000' },
      }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.error?.message).toContain('PowerPointApi 1.4');
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

/* ───────────────────────────── actuate: add-table-slide ───────────────────────────── */

function addTableSlide(params: ActuationRequest['params'], id = 'add-table-1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'add-table-slide', surface: 'powerpoint', params };
}

describe('PowerPointBridge.actuate add-table-slide', () => {
  it('reports uncertainty when a table landed but its identifying receipt could not be read', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const realSync = FakeContext.prototype.sync;
    const sync = vi
      .spyOn(FakeContext.prototype, 'sync')
      // The read-only lookups succeed; the sync that carries the table write fails.
      .mockImplementation(function (this: FakeContext) {
        return d.addedShapes.some((shape) => shape.kind === 'table')
          ? Promise.reject(new Error('Office disconnected'))
          : realSync.call(this);
      });
    try {
      const result = await new PowerPointBridge().actuate(
        addTableSlide({
          target: { slideId: 's1' },
          tableGrid: { hasHeaders: false, rows: [['Created']] },
        }),
      );
      expect(result).toMatchObject({
        ok: false,
        recoveryPending: true,
        error: { code: 'outcome_unknown' },
      });
      expect(d.slides[0]?.shapes[2]?.tableGrid).toEqual([['Created']]);
    } finally {
      sync.mockRestore();
    }
  });

  it('seeds a native table from the value grid and records the minted shape id as inverse', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addTableSlide(
        {
          target: { slideId: 's3' },
          tableGrid: {
            hasHeaders: true,
            rows: [
              ['Metric', 'Value'],
              ['ARR', '$12M'],
            ],
            left: 72,
            top: 120,
            width: 560,
            height: 260,
          },
        },
        'chg-table',
      ),
    );

    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-table'),
      kind: 'add-table-slide',
      location: 'shape:s3:s3-shape-2',
      inverse: { op: 'delete-object', objectType: 'shape', name: 's3-shape-2' },
    });
    expect(d.addedShapes).toEqual([
      {
        slideId: 's3',
        kind: 'table',
        detail: '2x2',
        options: { left: 72, top: 120, width: 560, height: 260 },
      },
    ]);
    expect(d.slides[2]?.shapes[2]?.tableGrid).toEqual([
      ['Metric', 'Value'],
      ['ARR', '$12M'],
    ]);
  });

  it('pads ragged rows out to the widest row length', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addTableSlide({
        target: { slideId: 's1' },
        tableGrid: { hasHeaders: false, rows: [['Only'], ['B', 'C']] },
      }),
    );

    expect(res.ok).toBe(true);
    expect(res.location).toBe('shape:s1:s1-shape-2');
    expect(d.addedShapes[0]).toMatchObject({ kind: 'table', detail: '2x2' });
    expect(d.slides[0]?.shapes[2]?.tableGrid).toEqual([
      ['Only', ''],
      ['B', 'C'],
    ]);
  });

  it('fails closed without touching the host when the target or grid is missing/empty', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const bridge = new PowerPointBridge();

    const noSlide = await bridge.actuate(
      addTableSlide({ tableGrid: { hasHeaders: false, rows: [['a']] } }),
    );
    expect(noSlide).toMatchObject({ ok: false, error: { code: 'no_target' } });

    const noGrid = await bridge.actuate(addTableSlide({ target: { slideId: 's1' } }));
    expect(noGrid).toMatchObject({ ok: false, error: { code: 'no_table' } });

    const emptyGrid = await bridge.actuate(
      addTableSlide({ target: { slideId: 's1' }, tableGrid: { hasHeaders: false, rows: [] } }),
    );
    expect(emptyGrid).toMatchObject({ ok: false, error: { code: 'no_table' } });

    const emptyRow = await bridge.actuate(
      addTableSlide({ target: { slideId: 's1' }, tableGrid: { hasHeaders: false, rows: [[]] } }),
    );
    expect(emptyRow).toMatchObject({ ok: false, error: { code: 'no_table' } });

    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('resolves slide=last and rejects an unknown slide before writing', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const ok = await new PowerPointBridge().actuate(
      addTableSlide({
        target: { slideId: 'last' },
        tableGrid: { hasHeaders: false, rows: [['a']] },
      }),
    );
    expect(ok.ok).toBe(true);
    expect(d.addedShapes[0]?.slideId).toBe('s3');

    const missing = await new PowerPointBridge().actuate(
      addTableSlide({ target: { slideId: 's9' }, tableGrid: { hasHeaders: false, rows: [['a']] } }),
    );
    expect(missing).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(d.addedShapes).toHaveLength(1);
  });

  it('degrades to unsupported below PowerPointApi 1.8 without touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.4 } });
    const res = await new PowerPointBridge().actuate(
      addTableSlide({ target: { slideId: 's1' }, tableGrid: { hasHeaders: false, rows: [['a']] } }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.error?.message).toContain('PowerPointApi 1.8');
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

/* ───────────────────────────── watch ───────────────────────────── */

describe('PowerPointBridge.watch', () => {
  it('registers selection + view handlers and emits the mapped local HostEvents', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const events: HostEvent[] = [];
    const unsub = new PowerPointBridge().watch((e) => events.push(e));

    // Two handlers registered: selection-changed and active-view-changed.
    expect(installed.office.added).toHaveLength(2);
    const [selReg, viewReg] = installed.office.added;
    selReg?.handler();
    viewReg?.handler();

    expect(events).toEqual([
      { type: 'selection-changed', surface: 'powerpoint', origin: 'local' },
      { type: 'document-changed', surface: 'powerpoint', origin: 'local' },
    ]);

    unsub();
    // Teardown removed both handlers.
    expect(installed.office.removed).toHaveLength(2);
  });

  it('is idempotent: a second unsubscribe does not remove handlers again', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const unsub = new PowerPointBridge().watch(() => {});
    unsub();
    const removedAfterFirst = installed.office.removed.length;
    unsub();
    expect(installed.office.removed.length).toBe(removedAfterFirst);
  });

  it('never throws when the host has no event bus (addHandlerAsync throws) and unsub is a no-op', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { officeThrowsOnAdd: true });
    const events: HostEvent[] = [];
    const bridge = new PowerPointBridge();
    let unsub: () => void = () => {};
    expect(() => {
      unsub = bridge.watch((e) => events.push(e));
    }).not.toThrow();
    expect(installed.office.added).toHaveLength(0);
    // Nothing was registered, so teardown removes nothing and does not throw.
    expect(() => unsub()).not.toThrow();
    expect(installed.office.removed).toHaveLength(0);
  });

  it('swallows a throwing removeHandlerAsync during teardown (best-effort)', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { officeThrowsOnRemove: true });
    const unsub = new PowerPointBridge().watch(() => {});
    expect(() => unsub()).not.toThrow();
  });
});
