import { afterEach, describe, expect, it, vi } from 'vitest';
import { asChangeId, contentHash } from '@ge/contracts';
import { captureNativeDeckTemplate, readNativeLayout } from './deck-template.js';
import { PowerPointBridge } from './powerpoint-bridge.js';

const load = (): void => undefined;
function fixture({
  shapeCount = 2,
  masters = 1,
  dropText = false,
  failAfterWrite = false,
  supported = true,
} = {}) {
  const shape = (id: string) => ({
    id,
    type: 'Placeholder',
    left: 10,
    top: 20,
    width: 200,
    height: 40,
    load,
    placeholderFormat: { type: 'Body', containedType: null, load },
    textFrame: { textRange: { text: `Manual ${id}`, load } },
  });
  const shapes = Array.from({ length: shapeCount }, (_, i) => shape(`shape-${i}`));
  const collection = (items: typeof shapes) => ({
    items,
    load,
    getCount: () => ({ value: items.length }),
    getItemAt: (index: number) => items[index],
  });
  const prior = { id: 'prior', name: 'Prior', type: 'Title', load, shapes: collection(shapes) };
  const next = {
    id: 'layout',
    name: 'Board pack',
    type: 'TitleOnly',
    load,
    shapes: collection(shapes),
  };
  let layout = prior;
  let applied = 0;
  const slide = {
    id: 'slide',
    load,
    shapes: collection(shapes),
    get layout() {
      return layout;
    },
    applyLayout: () => {
      applied++;
      layout = next;
      if (dropText) shapes[0]!.textFrame.textRange.text = 'lost';
    },
  };
  const ctx = {
    presentation: {
      slides: { items: [slide], load, getItem: () => slide },
      slideMasters: {
        getCount: () => ({ value: masters }),
        getItemAt: () => ({
          id: 'master',
          name: 'Brand',
          load,
          layouts: { getCount: () => ({ value: 1 }), getItemAt: () => next },
        }),
        getItem: () => ({ layouts: { getItem: () => next } }),
      },
    },
    sync: async () => {
      if (applied && failAfterWrite) throw new Error('connection lost');
    },
  };
  vi.stubGlobal('Office', {
    context: {
      document: { url: 'https://tenant/deck.pptx', mode: 'ReadWrite' },
      requirements: { isSetSupported: () => supported },
    },
    DocumentMode: { ReadWrite: 'ReadWrite', ReadOnly: 'ReadOnly' },
  });
  vi.stubGlobal('PowerPoint', {
    run: async (callback: (context: PowerPoint.RequestContext) => Promise<unknown>) =>
      callback(ctx as unknown as PowerPoint.RequestContext),
  });
  return {
    ctx: ctx as unknown as PowerPoint.RequestContext,
    next: next as unknown as PowerPoint.SlideLayout,
    applied: () => applied,
  };
}
afterEach(() => {
  vi.unstubAllGlobals();
});

describe('bounded native template capture', () => {
  it('captures native master/layout identities, geometry and hash without slide text', async () => {
    fixture();
    const result = await captureNativeDeckTemplate('https://tenant/deck.pptx');
    expect(result.complete).toBe(true);
    expect(result.masters[0]?.layouts[0]).toMatchObject({
      id: 'layout',
      name: 'Board pack',
      type: 'TitleOnly',
      revision: expect.stringMatching(/^sha256:/),
      placeholders: [
        { id: 'shape-0', type: 'Body', left: 10, top: 20, width: 200, height: 40 },
        { id: 'shape-1', type: 'Body', left: 10, top: 20, width: 200, height: 40 },
      ],
    });
    expect(JSON.stringify(result)).not.toContain('Manual');
  });
  it('bounds masters and shape metadata and discloses omissions', async () => {
    fixture({ shapeCount: 200, masters: 18 });
    const result = await captureNativeDeckTemplate('https://tenant/deck.pptx');
    expect(result.complete).toBe(false);
    expect(result.masters).toHaveLength(16);
    expect(
      result.masters.reduce(
        (sum, master) =>
          sum + master.layouts.reduce((count, layout) => count + layout.placeholders.length, 0),
        0,
      ),
    ).toBe(256);
    expect(result.omissions).toContain('2 additional masters were not captured.');
  });
  it('refuses template capture without its native requirement set', async () => {
    fixture({ supported: false });
    await expect(captureNativeDeckTemplate('deck')).rejects.toThrow('PowerPointApi 1.3');
  });
});

describe('native captured layout application', () => {
  async function request(h: ReturnType<typeof fixture>) {
    const capture = await readNativeLayout(h.ctx, h.next, 128);
    return {
      kind: 'apply-slide-layout',
      surface: 'powerpoint',
      changeId: asChangeId('layout-test'),
      params: {
        target: {
          slideId: 'slide',
          expectedDocumentId: await contentHash({
            surface: 'powerpoint',
            documentKey: 'https://tenant/deck.pptx',
          }),
        },
        layout: {
          masterId: 'master',
          layoutId: 'layout',
          expectedLayoutHash: capture.layout.revision,
        },
      },
    } as const;
  }
  it('uses the actual handler and verifies the layout and preservation of manual text', async () => {
    const h = fixture();
    const result = await new PowerPointBridge().actuate(await request(h));
    expect(result).toMatchObject({
      ok: true,
      completion: 'verified',
      verification: { status: 'verified' },
    });
    expect(h.applied()).toBe(1);
  });
  it('refuses a changed captured layout before mutation', async () => {
    const h = fixture();
    const r = await request(h);
    h.next.shapes.getItemAt(0).left = 99;
    expect(await new PowerPointBridge().actuate(r)).toMatchObject({
      ok: false,
      error: { code: 'target_conflict' },
    });
    expect(h.applied()).toBe(0);
  });
  it('reports unknown when layout application loses text or its sync is interrupted', async () => {
    for (const options of [{ dropText: true }, { failAfterWrite: true }]) {
      const h = fixture(options);
      const result = await new PowerPointBridge().actuate(await request(h));
      expect(
        result.verification?.status === 'unknown' || result.error?.code === 'outcome_unknown',
      ).toBe(true);
      expect(h.applied()).toBe(1);
    }
  });
});
