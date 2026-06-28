import { afterEach, describe, it, expect } from 'vitest';
import {
  asChangeId,
  DocStateSnapshotSchema,
  ResolvedContextSchema,
  type ActuationRequest,
  type ContextRef,
} from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { MAX_READ_SLIDES, PowerPointBridge } from './powerpoint-bridge.js';

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
}
interface SlideSeed {
  id: string;
  shapes: ShapeSeed[];
}
interface DeckSeed {
  slides: SlideSeed[];
  /** Zero-based indices of selected slides (the bridge reads items[0]). */
  selectedIndices: number[];
  selectedShapeIds: string[];
  insertedDecks: string[];
  insertedDeckOptions: PowerPoint.InsertSlideOptions[];
}

class FakeTextRange {
  constructor(private readonly shape: ShapeSeed) {}
  get text(): string {
    return this.shape.text;
  }
  set text(v: string) {
    this.shape.text = v;
  }
  load(_p?: string): this {
    return this;
  }
}

class FakeShape {
  readonly textFrame: { textRange: FakeTextRange };
  constructor(
    readonly shape: ShapeSeed,
    readonly id: string,
    readonly isNullObject = false,
  ) {
    this.textFrame = { textRange: new FakeTextRange(shape) };
  }
  load(_p?: string): this {
    return this;
  }
}

class FakeShapeCollection {
  items: FakeShape[];
  constructor(
    private readonly slide: SlideSeed,
    private readonly slideId: string,
  ) {
    this.items = slide.shapes.map((s, i) => new FakeShape(s, `${slide.id}-shape-${i}`));
  }
  load(_p?: string): this {
    return this;
  }
  getItemOrNullObject(id: string): FakeShape {
    const index = this.items.findIndex((shape) => shape.id === id);
    const existing = this.items[index];
    if (existing) return existing;
    return new FakeShape({ text: '' }, id || `${this.slideId}-missing-shape`, true);
  }
}

class FakeSlide {
  constructor(
    private readonly slide: SlideSeed,
    readonly index: number,
    private readonly seed: DeckSeed,
  ) {}
  get id(): string {
    return this.slide.id;
  }
  load(_p?: string): this {
    return this;
  }
  get shapes(): FakeShapeCollection {
    return new FakeShapeCollection(this.slide, this.slide.id);
  }
  setSelectedShapes(shapeIds: string[]): void {
    this.seed.selectedIndices = [this.index];
    this.seed.selectedShapeIds = [...shapeIds];
  }
}

class FakeSlideCollection {
  items: FakeSlide[];
  /** Recorded calls so tests can assert what the bridge drove. */
  addCalls = 0;
  constructor(private readonly seed: DeckSeed) {
    this.items = seed.slides.map((s, i) => new FakeSlide(s, i, seed));
  }
  load(_p?: string): this {
    return this;
  }
  getCount(): { value: number } {
    return { value: this.seed.slides.length };
  }
  add(): void {
    this.addCalls += 1;
    this.seed.slides.push({
      id: `sim-slide-${this.seed.slides.length + 1}`,
      shapes: [{ text: '' }, { text: '' }],
    });
    this.items = this.seed.slides.map((s, i) => new FakeSlide(s, i, this.seed));
  }
  getItemAt(index: number): FakeSlide {
    const slide = this.seed.slides[index];
    if (!slide) throw new Error(`fake-powerpoint: no slide at ${index}`);
    return new FakeSlide(slide, index, this.seed);
  }
  getItem(id: string): FakeSlide {
    const index = this.seed.slides.findIndex((s) => s.id === id);
    const slide = this.seed.slides[index];
    if (!slide) throw new Error(`fake-powerpoint: no slide ${id}`);
    return new FakeSlide(slide, index, this.seed);
  }
}

class FakePresentation {
  insertCalls: string[] = [];
  constructor(private readonly seed: DeckSeed) {}
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

class FakeContext {
  readonly presentation: FakePresentation;
  syncs = 0;
  constructor(seed: DeckSeed) {
    this.presentation = new FakePresentation(seed);
  }
  sync(): Promise<void> {
    this.syncs += 1;
    return Promise.resolve();
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

  const prevPP = (globalThis as Record<string, unknown>).PowerPoint;
  const prevOffice = (globalThis as Record<string, unknown>).Office;

  (globalThis as Record<string, unknown>).PowerPoint = {
    run: async <T>(cb: (ctx: FakeContext) => Promise<T>): Promise<T> => {
      const ctx = new FakeContext(seed);
      ctxRef.ctx = ctx;
      return cb(ctx);
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
          const v = version ? Number.parseFloat(version) : 0;
          return v <= max;
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
  const copy = slides.map((s) => ({ id: s.id, shapes: s.shapes.map((sh) => ({ text: sh.text })) }));
  return {
    slides: copy,
    selectedIndices,
    selectedShapeIds: [],
    insertedDecks: [],
    insertedDeckOptions: [],
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
  { id: 's1', shapes: [{ text: 'SLA Terms' }, { text: '99.5% contracted\nMonthly window' }] },
  { id: 's2', shapes: [{ text: 'Roster' }, { text: 'Pat, Sam' }] },
  { id: 's3', shapes: [{ text: 'Risk Summary' }, { text: 'SLA gap flagged' }] },
];

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
      title: 'Shape 1 on slide 2',
      preview: 'Roster',
    });
    expect(refs[2]).toMatchObject({
      id: 'pp:shape:s2:s2-shape-1',
      kind: 'shape',
      title: 'Shape 2 on slide 2',
      preview: 'Pat, Sam',
    });
    expect(refs[3]).toMatchObject({ id: 'pp:deck', kind: 'document' });
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
});

/* ───────────────────────────── captureDocState ───────────────────────────── */

describe('PowerPointBridge.captureDocState', () => {
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
      surface: 'excel',
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
    expect(appended?.shapes[1]?.text).toBe('Point A\nPoint B');
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

describe('PowerPointBridge.actuate set-shape-text', () => {
  it('rewrites one addressed shape and records the inverse text', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      setShapeText('s2', 's2-shape-1', 'Pat, Sam, and Lee'),
    );

    expect(res).toEqual({
      ok: true,
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
