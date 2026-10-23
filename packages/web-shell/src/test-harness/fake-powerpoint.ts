/**
 * In-memory **PowerPoint host simulator**. Models the slice of the Office.js object model the real
 * {@link "@ge/bridge-powerpoint"!PowerPointBridge} drives, so the REAL bridge runs unchanged against
 * a seeded deck. (PowerPoint exposes no object-model events in this typings, so its `watch()` uses
 * the Office bus — handled by the shared {@link "./fake-office"!makeFakeOffice} fake.)
 *
 * Enumerated host calls modelled (the fidelity boundary for PowerPoint):
 *   - `PowerPoint.run(cb)`.
 *   - `ctx.presentation.getSelectedSlides()` → items `{ id, index }`.
 *   - `ctx.presentation.slides` → items `{ id, index }`; `.getCount()` → `{ value }`;
 *     `.add()` (append a blank slide, WRITE); `.getItemAt(i)`; `.getItem(id)` (throws
 *     `ItemNotFound` for an unknown id).
 *   - `slide.shapes` → items; `slide.load('id,index')`; `slide.moveTo(index)` (WRITE).
 *   - `shapes.getItemAt(i)` / `.getItem(id)` (throws `ItemNotFound`) / `.getItemOrNullObject(id)`
 *     (`isNullObject: true` when absent).
 *   - `shapes.addTextBox(text, opts)` / `.addGeometricShape(type, opts)` / `.addLine(type, opts)` /
 *     `.addTable(rows, cols, opts)` (WRITEs; each new shape gets a stable minted id).
 *   - `shape.id/type/left/top/width/height`, `shape.placeholderFormat.type/containedType`,
 *     `shape.fill.foregroundColor`, `shape.lineFormat.color`, `shape.setZOrder(z)` (WRITE),
 *     `shape.getTable()` → `table.values` + `table.getCellOrNullObject(r, c).text` (read AND write).
 *   - `shape.textFrame.textRange` → `.text` (read AND write), `.font.bold/italic/underline/color/
 *     size/name` (read AND write), `.getSubstring(start, length).font` (the whole range's font).
 *   - `ctx.presentation.insertSlidesFromBase64(b64)` (prebuilt-deck WRITE).
 *   - `Office.context.document.addHandlerAsync(...)` for selection/view events (shared Office fake).
 *
 * Fidelity notes / boundary:
 *   - `load()` + `ctx.sync()` are not gated: every property reads through to the seed immediately
 *     (a real proxy would need the load first). Host errors (`ItemNotFound`, `InvalidArgument`) are
 *     thrown at call time rather than at the next `sync()`; inside `PowerPoint.run` the bridge sees
 *     the same rejection either way.
 *   - `shape.textFrame` throws `InvalidArgument` for a shape type with no text frame (Table, Line,
 *     Image, …) and `shape.getTable()` throws for a non-table shape, as the host does.
 *   - A shape without a seeded `id` is given `<slideId>-shape-<index>` the first time its slide's
 *     shapes are read, and keeps it (so ids stay stable when shapes are added or reordered).
 *   - A freshly `add()`ed slide is seeded with two empty placeholder shapes typed `Title` and `Body`
 *     (`Shape.type` 'Placeholder' + `placeholderFormat.type`), as a "Title and Content" layout gives.
 *     We do NOT model real slide layouts/masters — just enough shape structure for the native compose
 *     path to write into.
 *   - Font facets read `undefined` until written (no theme defaults are modelled).
 *   - `insertSlidesFromBase64` records the base64 payload and appends a marker slide; we do not parse
 *     PPTX bytes (out of scope — the bridge only needs the call to succeed).
 *   - Requirement sets come from the shared Office fake (default `PowerPointApi` 1.8), so 1.10-only
 *     paths (`presentation.pageSetup`) are skipped by the bridge and not modelled.
 */

import { installGlobal, composeRestores } from './globals.js';
import {
  makeFakeOffice,
  makeOfficeSeed,
  type OfficeSeed,
  type OfficeHandlerRegistry,
} from './fake-office.js';

/** A shape's text font facets (`TextRange.font`); unset facets read `undefined`. */
export interface ShapeFontSeed {
  bold?: boolean;
  italic?: boolean;
  /** `ShapeFont.underline` literal, e.g. 'Single' / 'None'. */
  underline?: string;
  color?: string;
  size?: number;
  name?: string;
}

/** A shape on a slide: its text frame's text plus the facets the bridge reads/writes. */
export interface ShapeSeed {
  text: string;
  /** `Shape.type`; defaults to 'TextBox'. */
  type?: string;
  /** `placeholderFormat.type` for a 'Placeholder' shape (e.g. 'Title', 'Body'). */
  placeholderType?: string;
  /** Stable host id; assigned `<slideId>-shape-<index>` on first read when absent. */
  id?: string;
  /** `addGeometricShape` geometry (e.g. 'Rectangle') for a 'GeometricShape'. */
  geometricShapeType?: string;
  /** `addLine` connector type (e.g. 'Straight') for a 'Line'. */
  connectorType?: string;
  /** `textFrame.textRange.paragraphFormat.bulletFormat.visible` (PowerPointApi 1.4), once set. */
  bulletVisible?: boolean;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  /** `shape.fill.foregroundColor`. */
  fill?: string;
  /** `shape.lineFormat.color`. */
  lineColor?: string;
  font?: ShapeFontSeed;
  /** Cell text grid of a 'Table' shape (`table.values`). */
  table?: string[][];
  /** Header-row fill from `TableAddOptions.specificCellProperties` (row 0), as the host was given. */
  tableHeaderFill?: string;
}

/** A slide: a stable id + its shapes (title shape first, by layout convention). */
export interface SlideSeed {
  id: string;
  shapes: ShapeSeed[];
}

/** The deck seed: ordered slides + the indices currently selected. */
export interface PowerPointSeed {
  slides: SlideSeed[];
  /** Zero-based indices of the selected slides (the bridge reads `items[0]`). */
  selectedIndices: number[];
  /** Recorded base64 decks merged via `insertSlidesFromBase64`. */
  insertedDecks: string[];
}

/* ─────────────────────────── fake object model ─────────────────────────── */

/** Shape types with a text frame (`Shape.textFrame` throws `InvalidArgument` for the rest). */
const TEXT_FRAME_TYPES: ReadonlySet<string> = new Set([
  'GeometricShape',
  'TextBox',
  'Placeholder',
  'Callout',
  'Freeform',
]);

/** An Office.js-style host error (`OfficeExtension.Error` carries a string `code`). */
function hostError(code: string, message: string): Error {
  return Object.assign(new Error(`fake-powerpoint: ${message}`), { code });
}

function shapeIdTaken(slide: SlideSeed, id: string): boolean {
  return slide.shapes.some((shape) => shape.id === id);
}

/** Give every shape on `slide` a stable id (`<slideId>-shape-<index>` unless taken). */
function ensureShapeIds(slide: SlideSeed): void {
  slide.shapes.forEach((shape, index) => {
    if (shape.id !== undefined) return;
    shape.id = shapeIdTaken(slide, `${slide.id}-shape-${index}`)
      ? mintShapeId(slide)
      : `${slide.id}-shape-${index}`;
  });
}

/** A fresh id for a new shape on `slide`, never reusing one already on it. */
function mintShapeId(slide: SlideSeed): string {
  let n = slide.shapes.length;
  while (shapeIdTaken(slide, `${slide.id}-shape-${n}`)) n += 1;
  return `${slide.id}-shape-${n}`;
}

class FakeShapeFont {
  constructor(private readonly shape: ShapeSeed) {}
  private get facets(): ShapeFontSeed {
    this.shape.font ??= {};
    return this.shape.font;
  }
  load(_props?: string): this {
    return this;
  }
  get bold(): boolean | undefined {
    return this.shape.font?.bold;
  }
  set bold(value: boolean | undefined) {
    this.facets.bold = value;
  }
  get italic(): boolean | undefined {
    return this.shape.font?.italic;
  }
  set italic(value: boolean | undefined) {
    this.facets.italic = value;
  }
  get underline(): string | undefined {
    return this.shape.font?.underline;
  }
  set underline(value: string | undefined) {
    this.facets.underline = value;
  }
  get color(): string | undefined {
    return this.shape.font?.color;
  }
  set color(value: string | undefined) {
    this.facets.color = value;
  }
  get size(): number | undefined {
    return this.shape.font?.size;
  }
  set size(value: number | undefined) {
    this.facets.size = value;
  }
  get name(): string | undefined {
    return this.shape.font?.name;
  }
  set name(value: string | undefined) {
    this.facets.name = value;
  }
}

class FakeTextRange {
  readonly font: FakeShapeFont;
  /** PowerPointApi 1.4: whether the paragraphs show the layout's bullet. */
  readonly paragraphFormat: { bulletFormat: { visible: boolean | undefined } };
  constructor(private readonly shape: ShapeSeed) {
    this.font = new FakeShapeFont(shape);
    this.paragraphFormat = {
      bulletFormat: {
        get visible(): boolean | undefined {
          return shape.bulletVisible;
        },
        set visible(value: boolean | undefined) {
          shape.bulletVisible = value;
        },
      },
    };
  }
  get text(): string {
    return this.shape.text;
  }
  set text(value: string) {
    this.shape.text = value;
  }
  load(_props?: string): this {
    return this;
  }
  /** A sub-range; its font is the whole range's (per-run formatting is not modelled). */
  getSubstring(_start: number, _length?: number): { font: FakeShapeFont; load(): unknown } {
    const sub = { font: this.font, load: () => sub };
    return sub;
  }
}

class FakeTableCell {
  readonly isNullObject = false;
  constructor(
    private readonly values: string[][],
    private readonly row: number,
    private readonly column: number,
  ) {}
  load(_props?: string): this {
    return this;
  }
  get text(): string {
    return this.values[this.row]?.[this.column] ?? '';
  }
  set text(value: string) {
    const row = this.values[this.row];
    if (row) row[this.column] = value;
  }
}

/** `getCellOrNullObject` outside the grid: a null object; writes to it are dropped. */
const NULL_TABLE_CELL = {
  isNullObject: true,
  text: '',
  load(): unknown {
    return NULL_TABLE_CELL;
  },
};

class FakeTable {
  constructor(private readonly grid: string[][]) {}
  load(_props?: string): this {
    return this;
  }
  get values(): string[][] {
    return this.grid.map((row) => [...row]);
  }
  get rowCount(): number {
    return this.grid.length;
  }
  get columnCount(): number {
    return this.grid[0]?.length ?? 0;
  }
  getCellOrNullObject(row: number, column: number): FakeTableCell | typeof NULL_TABLE_CELL {
    if (row < 0 || column < 0 || row >= this.rowCount || column >= this.columnCount) {
      return { ...NULL_TABLE_CELL };
    }
    return new FakeTableCell(this.grid, row, column);
  }
}

class FakeShape {
  readonly isNullObject = false;
  constructor(
    private readonly slide: SlideSeed,
    readonly shape: ShapeSeed,
  ) {}
  load(_props?: string): this {
    return this;
  }
  get id(): string {
    return this.shape.id ?? '';
  }
  get type(): string {
    return this.shape.type ?? 'TextBox';
  }
  get left(): number | undefined {
    return this.shape.left;
  }
  get top(): number | undefined {
    return this.shape.top;
  }
  get width(): number | undefined {
    return this.shape.width;
  }
  get height(): number | undefined {
    return this.shape.height;
  }
  get textFrame(): { textRange: FakeTextRange } {
    if (!TEXT_FRAME_TYPES.has(this.type)) {
      throw hostError('InvalidArgument', `a ${this.type} shape has no text frame`);
    }
    return { textRange: new FakeTextRange(this.shape) };
  }
  get placeholderFormat(): { type: string; containedType: null; load(_props?: string): unknown } {
    const format = {
      type: this.shape.placeholderType ?? 'Unsupported',
      containedType: null,
      load: () => format,
    };
    return format;
  }
  get fill(): {
    foregroundColor: string | undefined;
    load(_props?: string): unknown;
    setSolidColor(color: string): void;
  } {
    const shape = this.shape;
    const fill = {
      get foregroundColor(): string | undefined {
        return shape.fill;
      },
      set foregroundColor(value: string | undefined) {
        shape.fill = value;
      },
      load: () => fill,
      setSolidColor: (color: string) => {
        shape.fill = color;
      },
    };
    return fill;
  }
  get lineFormat(): { color: string | undefined; load(_props?: string): unknown } {
    const shape = this.shape;
    const line = {
      get color(): string | undefined {
        return shape.lineColor;
      },
      set color(value: string | undefined) {
        shape.lineColor = value;
      },
      load: () => line,
    };
    return line;
  }
  getTable(): FakeTable {
    if (this.type !== 'Table' || !this.shape.table) {
      throw hostError('InvalidArgument', `a ${this.type} shape is not a table`);
    }
    return new FakeTable(this.shape.table);
  }
  /** Reorder within the slide's shape list (list order is z-order, back to front). */
  setZOrder(position: string): void {
    const shapes = this.slide.shapes;
    const from = shapes.indexOf(this.shape);
    if (from < 0) return;
    const to =
      position === 'BringToFront'
        ? shapes.length - 1
        : position === 'SendToBack'
          ? 0
          : position === 'BringForward'
            ? Math.min(shapes.length - 1, from + 1)
            : position === 'SendBackward'
              ? Math.max(0, from - 1)
              : from;
    shapes.splice(from, 1);
    shapes.splice(to, 0, this.shape);
  }
}

/** `getItemOrNullObject` for an absent id: only `isNullObject` is meaningful. */
class FakeNullShape {
  readonly isNullObject = true;
  load(_props?: string): this {
    return this;
  }
}

/** Geometry accepted by the `add*` calls (`ShapeAddOptions` / `TableAddOptions`). */
interface FakeAddOptions {
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  specificCellProperties?: Array<Array<{ fill?: { color?: string } }>>;
}

class FakeShapeCollection {
  items: FakeShape[];
  constructor(private readonly slide: SlideSeed) {
    ensureShapeIds(slide);
    this.items = slide.shapes.map((s) => new FakeShape(slide, s));
  }
  load(_props?: string): this {
    return this;
  }
  getCount(): { value: number } {
    return { value: this.slide.shapes.length };
  }
  getItemAt(index: number): FakeShape {
    const shape = this.slide.shapes[index];
    if (!shape) throw hostError('ItemNotFound', `no shape at index ${index}`);
    return new FakeShape(this.slide, shape);
  }
  getItem(id: string): FakeShape {
    const shape = this.slide.shapes.find((s) => s.id === id);
    if (!shape) throw hostError('ItemNotFound', `no shape ${id}`);
    return new FakeShape(this.slide, shape);
  }
  getItemOrNullObject(id: string): FakeShape | FakeNullShape {
    const shape = this.slide.shapes.find((s) => s.id === id);
    return shape ? new FakeShape(this.slide, shape) : new FakeNullShape();
  }
  addTextBox(text: string, options?: FakeAddOptions): FakeShape {
    return this.append({ text, type: 'TextBox', ...geometry(options) });
  }
  addGeometricShape(geometricShapeType: string, options?: FakeAddOptions): FakeShape {
    return this.append({
      text: '',
      type: 'GeometricShape',
      geometricShapeType,
      ...geometry(options),
    });
  }
  addLine(connectorType: string = 'Straight', options?: FakeAddOptions): FakeShape {
    return this.append({ text: '', type: 'Line', connectorType, ...geometry(options) });
  }
  addTable(rowCount: number, columnCount: number, options?: FakeAddOptions): FakeShape {
    if (rowCount < 1 || columnCount < 1) {
      throw hostError('InvalidArgument', `a table needs at least one row and column`);
    }
    const table = Array.from({ length: rowCount }, () =>
      Array.from({ length: columnCount }, () => ''),
    );
    const headerFill = options?.specificCellProperties?.[0]?.[0]?.fill?.color;
    return this.append({
      text: '',
      type: 'Table',
      table,
      ...geometry(options),
      ...(headerFill ? { tableHeaderFill: headerFill } : {}),
    });
  }
  private append(shape: ShapeSeed): FakeShape {
    shape.id = mintShapeId(this.slide);
    this.slide.shapes.push(shape);
    const added = new FakeShape(this.slide, shape);
    this.items = [...this.items, added];
    return added;
  }
}

/** Copy only the geometry fields that were given (absent → host default, left unset here). */
function geometry(options: FakeAddOptions | undefined): FakeAddOptions {
  const out: FakeAddOptions = {};
  if (options?.left !== undefined) out.left = options.left;
  if (options?.top !== undefined) out.top = options.top;
  if (options?.width !== undefined) out.width = options.width;
  if (options?.height !== undefined) out.height = options.height;
  return out;
}

class FakeSlide {
  constructor(
    private readonly seed: PowerPointSeed,
    private readonly slide: SlideSeed,
    readonly index: number,
  ) {}
  get id(): string {
    return this.slide.id;
  }
  load(_props?: string): this {
    return this;
  }
  get shapes(): FakeShapeCollection {
    return new FakeShapeCollection(this.slide);
  }
  /** Move this slide to zero-based `index` (clamped to the deck). */
  moveTo(index: number): void {
    const slides = this.seed.slides;
    const from = slides.indexOf(this.slide);
    if (from < 0) throw hostError('ItemNotFound', `slide ${this.slide.id} is gone`);
    slides.splice(from, 1);
    slides.splice(Math.max(0, Math.min(index, slides.length)), 0, this.slide);
  }
}

class FakeSlideCollection {
  items: FakeSlide[];
  constructor(private readonly seed: PowerPointSeed) {
    this.items = seed.slides.map((s, i) => new FakeSlide(seed, s, i));
  }
  load(_props?: string): this {
    return this;
  }
  getCount(): { value: number } {
    return { value: this.seed.slides.length };
  }
  add(): void {
    // A new slide from a "Title and Content"-style layout: a title and a body placeholder.
    this.seed.slides.push({
      id: `sim-slide-${this.seed.slides.length + 1}`,
      shapes: [
        { text: '', type: 'Placeholder', placeholderType: 'Title' },
        { text: '', type: 'Placeholder', placeholderType: 'Body' },
      ],
    });
    this.items = this.seed.slides.map((s, i) => new FakeSlide(this.seed, s, i));
  }
  getItemAt(index: number): FakeSlide {
    const slide = this.seed.slides[index];
    if (!slide) throw hostError('ItemNotFound', `no slide at index ${index}`);
    return new FakeSlide(this.seed, slide, index);
  }
  getItem(id: string): FakeSlide {
    const index = this.seed.slides.findIndex((s) => s.id === id);
    const slide = this.seed.slides[index];
    if (!slide) throw hostError('ItemNotFound', `no slide ${id}`);
    return new FakeSlide(this.seed, slide, index);
  }
}

class FakePresentation {
  constructor(private readonly seed: PowerPointSeed) {}
  get slides(): FakeSlideCollection {
    return new FakeSlideCollection(this.seed);
  }
  getSelectedSlides(): FakeSlideCollection {
    const all = new FakeSlideCollection(this.seed);
    const sel = new FakeSlideCollection(this.seed);
    sel.items = this.seed.selectedIndices
      .map((i) => all.items[i])
      .filter((s): s is FakeSlide => s !== undefined);
    return sel;
  }
  insertSlidesFromBase64(base64: string): void {
    this.seed.insertedDecks.push(base64);
    this.seed.slides.push({
      id: `sim-inserted-${this.seed.insertedDecks.length}`,
      shapes: [{ text: '(inserted deck)' }],
    });
  }
}

class FakePowerPointContext {
  readonly presentation: FakePresentation;
  constructor(seed: PowerPointSeed) {
    this.presentation = new FakePresentation(seed);
  }
  sync(): Promise<void> {
    return Promise.resolve();
  }
}

/** The `PowerPoint` namespace object installed onto `globalThis.PowerPoint`. */
interface FakePowerPointNamespace {
  run<T>(callback: (ctx: FakePowerPointContext) => Promise<T>): Promise<T>;
}

/* ─────────────────────────── the simulator facade ──────────────────────── */

/** A shape as read back: its id + type always set, other facets only when present. */
export type ShapeSnapshot = Readonly<ShapeSeed & { id: string; type: string }>;

/** A read-back view of the deck after a run. */
export interface PowerPointSnapshot {
  slides: ReadonlyArray<{ id: string; shapeTexts: string[]; shapes: ShapeSnapshot[] }>;
  insertedDecks: ReadonlyArray<string>;
}

/** The installed PowerPoint simulator. */
export interface PowerPointSimulator {
  readonly seed: PowerPointSeed;
  readonly office: OfficeSeed;
  readonly officeHandlers: OfficeHandlerRegistry;
  snapshot(): PowerPointSnapshot;
  restore(): void;
}

/**
 * Install an in-memory PowerPoint host: writes `globalThis.PowerPoint` + `globalThis.Office` so the
 * REAL {@link "@ge/bridge-powerpoint"!PowerPointBridge} runs against `seed`. Defaults to
 * {@link defaultPowerPointSeed} + a modern requirement set.
 */
export function installFakePowerPoint(
  seed: PowerPointSeed = defaultPowerPointSeed(),
  requirements: Record<string, number> = { PowerPointApi: 8 },
): PowerPointSimulator {
  const office = makeOfficeSeed(requirements);
  const { office: officeNs, handlers: officeHandlers } = makeFakeOffice(office);

  const powerpoint: FakePowerPointNamespace = {
    run: async <T>(callback: (ctx: FakePowerPointContext) => Promise<T>): Promise<T> =>
      callback(new FakePowerPointContext(seed)),
  };

  const restore = composeRestores([
    installGlobal('PowerPoint', powerpoint),
    installGlobal('Office', officeNs),
  ]);

  return {
    seed,
    office,
    officeHandlers,
    snapshot: () => ({
      slides: seed.slides.map((s) => {
        ensureShapeIds(s);
        return {
          id: s.id,
          shapeTexts: s.shapes.map((sh) => sh.text),
          shapes: s.shapes.map(snapshotShape),
        };
      }),
      insertedDecks: [...seed.insertedDecks],
    }),
    restore,
  };
}

/** A detached copy of a shape seed (so a snapshot never aliases live deck state). */
function snapshotShape(shape: ShapeSeed): ShapeSnapshot {
  return {
    ...shape,
    id: shape.id ?? '',
    type: shape.type ?? 'TextBox',
    ...(shape.font ? { font: { ...shape.font } } : {}),
    ...(shape.table ? { table: shape.table.map((row) => [...row]) } : {}),
  };
}

/* ─────────────────────────── builders + default fixture ─────────────────── */

/** Build a {@link PowerPointSeed} from slides; defaults to selecting the first slide. */
export function powerPointSeed(init: {
  slides: SlideSeed[];
  selectedIndices?: number[];
}): PowerPointSeed {
  return {
    slides: init.slides,
    selectedIndices: init.selectedIndices ?? (init.slides.length > 0 ? [0] : []),
    insertedDecks: [],
  };
}

/** A realistic deck fixture: a title slide + two content slides (title shape + body shape each). */
export function defaultPowerPointSeed(): PowerPointSeed {
  return powerPointSeed({
    slides: [
      { id: 'slide-1', shapes: [{ text: 'Q3 Business Review' }, { text: 'Acme Corp · FY26' }] },
      {
        id: 'slide-2',
        shapes: [{ text: 'Revenue' }, { text: 'Revenue up 12% QoQ\nEast region leads growth' }],
      },
      {
        id: 'slide-3',
        shapes: [{ text: 'Risks' }, { text: 'Churn in West region\nSupply constraints in Q4' }],
      },
    ],
    selectedIndices: [1],
  });
}
