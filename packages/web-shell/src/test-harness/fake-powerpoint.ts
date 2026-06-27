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
 *     `.add()` (append a blank slide, WRITE); `.getItemAt(i)`.
 *   - `slide.shapes` → items; `slide.load('id,index')`.
 *   - `shape.textFrame.textRange` → `.text` (read AND write — the compose-slide WRITE).
 *   - `ctx.presentation.insertSlidesFromBase64(b64)` (prebuilt-deck WRITE).
 *   - `Office.context.document.addHandlerAsync(...)` for selection/view events (shared Office fake).
 *
 * Fidelity notes / boundary:
 *   - A freshly `add()`ed slide is seeded with two empty placeholder shapes (title + body), matching
 *     the bridge's layout convention (`shapes.items[0]` = title, `[1]` = body). We do NOT model real
 *     slide layouts/masters — just enough shape structure for the native compose path to write into.
 *   - `insertSlidesFromBase64` records the base64 payload and appends a marker slide; we do not parse
 *     PPTX bytes (out of scope — the bridge only needs the call to succeed).
 */

import { installGlobal, composeRestores } from './globals.js';
import {
  makeFakeOffice,
  makeOfficeSeed,
  type OfficeSeed,
  type OfficeHandlerRegistry,
} from './fake-office.js';

/** A shape on a slide: just its text frame's text (the only facet the bridge reads/writes). */
export interface ShapeSeed {
  text: string;
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

class FakeTextRange {
  constructor(
    private readonly shape: ShapeSeed,
    private _loaded = false,
  ) {}
  get text(): string {
    return this.shape.text;
  }
  set text(value: string) {
    this.shape.text = value;
  }
  load(_props?: string): this {
    this._loaded = true;
    return this;
  }
}

class FakeShape {
  readonly textFrame: { textRange: FakeTextRange };
  constructor(
    readonly shape: ShapeSeed,
    readonly id: string,
  ) {
    this.textFrame = { textRange: new FakeTextRange(shape) };
  }
}

class FakeShapeCollection {
  items: FakeShape[];
  constructor(slide: SlideSeed) {
    this.items = slide.shapes.map((s, i) => new FakeShape(s, `${slide.id}-shape-${i}`));
  }
  load(_props?: string): this {
    return this;
  }
}

class FakeSlide {
  constructor(
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
}

class FakeSlideCollection {
  items: FakeSlide[];
  constructor(private readonly seed: PowerPointSeed) {
    this.items = seed.slides.map((s, i) => new FakeSlide(s, i));
  }
  load(_props?: string): this {
    return this;
  }
  getCount(): { value: number } {
    return { value: this.seed.slides.length };
  }
  add(): void {
    // A blank slide with two placeholder shapes (title + body), matching the bridge's convention.
    this.seed.slides.push({
      id: `sim-slide-${this.seed.slides.length + 1}`,
      shapes: [{ text: '' }, { text: '' }],
    });
    this.items = this.seed.slides.map((s, i) => new FakeSlide(s, i));
  }
  getItemAt(index: number): FakeSlide {
    const slide = this.seed.slides[index];
    if (!slide) throw new Error(`fake-powerpoint: no slide at index ${index}`);
    return new FakeSlide(slide, index);
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

/** A read-back view of the deck after a run. */
export interface PowerPointSnapshot {
  slides: ReadonlyArray<{ id: string; shapeTexts: string[] }>;
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
      slides: seed.slides.map((s) => ({ id: s.id, shapeTexts: s.shapes.map((sh) => sh.text) })),
      insertedDecks: [...seed.insertedDecks],
    }),
    restore,
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
