import { describe, it, expect, afterEach } from 'vitest';
import { installFakeExcel, excelSeed } from './fake-excel.js';
import { installFakeWord, wordSeed } from './fake-word.js';
import { installFakePowerPoint } from './fake-powerpoint.js';

/**
 * Fidelity tests for the simulators THEMSELVES (no bridge in the loop): they assert the fakes model
 * the load/sync contract, record writes, re-resolve searches, and fire host events the way the real
 * Office.js seam does — so an integration test can't pass against a fake that diverges from the host.
 *
 * These probe the globals the bridges drive (`Excel.run` / `Word.run` / `PowerPoint.run`) directly,
 * cast through `unknown` only at the read boundary (mirroring how the bridges reach the typed global).
 */

let restore: (() => void) | undefined;
afterEach(() => {
  restore?.();
  restore = undefined;
});

/** Read a global host namespace the way a bridge would (the typed-global boundary). */
function host<T>(name: string): T {
  return (globalThis as unknown as Record<string, unknown>)[name] as T;
}

interface ExcelLike {
  run<T>(cb: (ctx: ExcelCtx) => Promise<T>): Promise<T>;
}
interface ExcelCtx {
  workbook: {
    worksheets: { getActiveWorksheet(): ExcelSheet; getItem(n: string): ExcelSheet };
    getSelectedRange(): ExcelRange;
    names: { load(p?: string): { items: Array<{ name: string; formula: string; type: string }> } };
  };
  sync(): Promise<void>;
}
interface ExcelSheet {
  load(p?: string): ExcelSheet;
  name: string;
  getUsedRange(): ExcelRange;
  getUsedRangeOrNullObject(): ExcelRange;
  getRange(a1: string): ExcelRange;
}
interface ExcelRange {
  load(p?: string): ExcelRange;
  address: string;
  values: string[][];
  isNullObject: boolean;
  rowCount: number;
  columnCount: number;
  getCell(r: number, c: number): ExcelRange;
}

describe('fake-excel fidelity', () => {
  it('reading a Range property before load()+sync() throws (load/sync fidelity)', async () => {
    const sim = installFakeExcel();
    restore = sim.restore;
    const Excel = host<ExcelLike>('Excel');

    await Excel.run(async (ctx) => {
      const used = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange();
      // No load() yet: reading .values must throw like the real proxy (PropertyNotLoaded).
      expect(() => used.values).toThrow(/not loaded/);
      // Loaded but not yet synced: still unreadable until context.sync() resolves it.
      used.load('values');
      expect(() => used.values).toThrow(/not loaded/);
      await ctx.sync();
      // After sync the loaded property is readable; an UNloaded sibling still throws.
      expect(used.values[0]).toEqual(['region', 'rep', 'revenue', 'cost']);
      expect(() => used.address).toThrow(/not loaded/);
    });
  });

  it('sync() resolves loads; used range + values read from the seed', async () => {
    const sim = installFakeExcel();
    restore = sim.restore;
    const Excel = host<ExcelLike>('Excel');

    const out = await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      const used = sheet.getUsedRange();
      used.load('address,values');
      await ctx.sync();
      return { name: sheet.name, address: used.address, header: used.values[0] };
    });
    expect(out.name).toBe('Sales');
    expect(out.address).toBe('Sales!A1:D7');
    expect(out.header).toEqual(['region', 'rep', 'revenue', 'cost']);
  });

  it('a queued values write commits into the seed at sync() (recorded for assertions)', async () => {
    const sim = installFakeExcel();
    restore = sim.restore;
    const Excel = host<ExcelLike>('Excel');

    await Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getItem('Summary').getRange('B2');
      range.values = [['written']];
      range.load('address');
      await ctx.sync();
    });
    const summary = sim.snapshot().sheets.find((s) => s.name === 'Summary');
    expect(summary?.values[1]?.[1]).toBe('written');
  });

  it('getUsedRangeOrNullObject() returns a null object for an empty sheet', async () => {
    const sim = installFakeExcel(
      excelSeed({ sheets: [{ name: 'Blank', origin: 'A1', values: [['']] }] }),
    );
    restore = sim.restore;
    const Excel = host<ExcelLike>('Excel');
    const isNull = await Excel.run(async (ctx) => {
      const used = ctx.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject();
      used.load('isNullObject');
      await ctx.sync();
      return used.isNullObject;
    });
    expect(isNull).toBe(true);
  });

  it('isSet() reads the seeded requirement set off the installed Office global', async () => {
    const sim = installFakeExcel(undefined, { ExcelApi: 9 });
    restore = sim.restore;
    const reqs = (
      globalThis as unknown as {
        Office: { context: { requirements: { isSetSupported(n: string, v: string): boolean } } };
      }
    ).Office.context.requirements;
    expect(reqs.isSetSupported('ExcelApi', '1.9')).toBe(true);
    expect(reqs.isSetSupported('ExcelApi', '1.10')).toBe(false);
  });

  it('selection / comment-add events fire the sinks a watcher registered', async () => {
    const sim = installFakeExcel();
    restore = sim.restore;
    const Excel = host<ExcelLike>('Excel');
    const seen: string[] = [];
    // Register on the same workbook collections the bridge's watch() uses.
    await Excel.run(async (ctx) => {
      const wb = ctx.workbook as unknown as {
        worksheets: { onSelectionChanged: { add(h: (a: { address: string }) => void): void } };
        comments: {
          onAdded: { add(h: (a: { commentDetails: Array<{ commentId: string }> }) => void): void };
        };
      };
      wb.worksheets.onSelectionChanged.add((a) => seen.push(`sel:${a.address}`));
      wb.comments.onAdded.add((a) => seen.push(`cmt:${a.commentDetails[0]?.commentId}`));
      await ctx.sync();
    });
    sim.events.fireSelectionChanged('Sales!B2');
    sim.events.fireCommentAdded('c-9');
    expect(seen).toEqual(['sel:Sales!B2', 'cmt:c-9']);
  });
});

interface WordLike {
  run<T>(cb: (ctx: WordCtx) => Promise<T>): Promise<T>;
  InsertLocation: { replace: string };
}
interface WordCtx {
  document: {
    body: {
      search(q: string, o?: { matchCase?: boolean }): { load(p?: string): void; items: WordHit[] };
    };
  };
  sync(): Promise<void>;
}
interface WordHit {
  text: string;
  insertText(t: string, loc: string): void;
}

describe('fake-word fidelity', () => {
  it('body.search re-resolves against the live body; insertText replaces in place', async () => {
    const sim = installFakeWord();
    restore = sim.restore;
    const Word = host<WordLike>('Word');

    // First search finds the anchor; after insertText replaces it, a re-search finds nothing —
    // the apply-time re-resolution the bridge's drift-degradation depends on.
    const counts = await Word.run(async (ctx) => {
      const before = ctx.document.body.search('99.5%');
      before.load('items');
      await ctx.sync();
      const beforeCount = before.items.length;
      before.items[0]?.insertText('replaced', Word.InsertLocation.replace);
      await ctx.sync();
      const after = ctx.document.body.search('99.5%');
      after.load('items');
      await ctx.sync();
      return { beforeCount, afterCount: after.items.length };
    });
    expect(counts.beforeCount).toBeGreaterThan(0);
    expect(counts.afterCount).toBe(0);
    expect(sim.snapshot().bodyText).toContain('replaced');
  });

  it('an absent anchor yields zero hits (the drift path)', async () => {
    const sim = installFakeWord(
      wordSeed({ paragraphs: [{ text: 'no service levels here', styleBuiltIn: 'Normal' }] }),
    );
    restore = sim.restore;
    const Word = host<WordLike>('Word');
    const n = await Word.run(async (ctx) => {
      const r = ctx.document.body.search('99.5%');
      r.load('items');
      await ctx.sync();
      return r.items.length;
    });
    expect(n).toBe(0);
  });
});

interface PowerPointLike {
  run<T>(cb: (ctx: PpCtx) => Promise<T>): Promise<T>;
}
interface PpCtx {
  presentation: {
    slides: { load(p?: string): void; items: Array<{ id: string; index: number }> };
    getSelectedSlides(): { load(p?: string): void; items: Array<{ id: string; index: number }> };
  };
  sync(): Promise<void>;
}

describe('fake-powerpoint fidelity', () => {
  it('the deck + selection read from the seed', async () => {
    const sim = installFakePowerPoint();
    restore = sim.restore;
    const PowerPoint = host<PowerPointLike>('PowerPoint');
    const out = await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id,items/index');
      const sel = ctx.presentation.getSelectedSlides();
      sel.load('items/index');
      await ctx.sync();
      return { count: slides.items.length, selectedIndex: sel.items[0]?.index };
    });
    expect(out.count).toBe(3);
    expect(out.selectedIndex).toBe(1);
  });
});
