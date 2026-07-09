// packages/runtime/src/docfs/doc-mount.test.ts
import { describe, expect, it } from 'vitest';
import type { DocBridge } from '../bridge.js';
import { docMount } from './doc-mount.js';

function fakeBridge(): DocBridge {
  return {
    surface: 'excel',
    async captureDocState() {
      return {
        surface: 'excel',
        version: 1,
        capturedAt: new Date(0).toISOString(),
        outline: [{ level: 1, text: 'Summary' }],
        inventory: [{ kind: 'sheet', id: 'sheet:Q3', title: 'Q3', summary: '9x4' }],
        selection: { kind: 'range', title: 'A1:D9', preview: 'rev…' },
      };
    },
    async searchDocument(q: string) {
      return q === 'revenue'
        ? [
            {
              ref: { id: 'r1', kind: 'range', surface: 'excel', title: 'A1' },
              value: { as: 'text', text: 'revenue 42' },
            },
          ]
        : [];
    },
  } as unknown as DocBridge;
}

function bridgeWithoutCapture(): DocBridge {
  return {
    surface: 'word',
  } as unknown as DocBridge;
}

describe('docMount', () => {
  it('lists the doc-state index (inventory + selection + outline)', async () => {
    const entries = await docMount(fakeBridge()).readdir('');
    const names = entries.map((e) => e.name);
    expect(names).toContain('outline.md');
    expect(names).toContain('selection');
    expect(names).toContain('sheet:Q3');
  });

  it('reads the outline as rendered markdown headings', async () => {
    const v = await docMount(fakeBridge()).readFile('outline.md');
    expect(v?.text).toContain('# Summary');
  });

  it('search delegates to searchDocument and renders a text-variant ResolvedContext', async () => {
    const hits = await docMount(fakeBridge()).search('', 'revenue');
    expect(hits[0]?.text).toContain('revenue 42');
    expect(hits[0]?.path).toBe('r1');
  });

  it('returns [] from search when searchDocument does not match', async () => {
    const hits = await docMount(fakeBridge()).search('', 'nope');
    expect(hits).toEqual([]);
  });

  it('reads the selection preview', async () => {
    const v = await docMount(fakeBridge()).readFile('selection');
    expect(v?.text).toBe('rev…');
  });

  it('reads an inventory entry via readRange', async () => {
    const bridge = {
      surface: 'excel',
      async captureDocState() {
        return {
          surface: 'excel',
          version: 1,
          capturedAt: new Date(0).toISOString(),
          outline: [],
          inventory: [{ kind: 'sheet', id: 'sheet:Q3', title: 'Q3' }],
        };
      },
      async readRange(a1: string) {
        expect(a1).toBe('sheet:Q3');
        return [
          {
            ref: { id: 'sheet:Q3', kind: 'sheet', surface: 'excel', title: 'Q3' },
            value: { as: 'text', text: 'Q3 numbers' },
          },
        ];
      },
    } as unknown as DocBridge;
    const v = await docMount(bridge).readFile('sheet:Q3');
    expect(v?.text).toContain('Q3 numbers');
  });

  it('degrades gracefully when captureDocState is absent: readdir returns just outline.md', async () => {
    const entries = await docMount(bridgeWithoutCapture()).readdir('');
    expect(entries.map((e) => e.name)).toEqual(['outline.md']);
  });

  it('degrades gracefully when captureDocState is absent: readFile("outline.md") does not throw', async () => {
    const v = await docMount(bridgeWithoutCapture()).readFile('outline.md');
    expect(v).not.toBeNull();
    expect(v?.text).toBe('');
  });

  it('returns null reading an inventory-id-shaped rel when captureDocState is absent', async () => {
    const v = await docMount(bridgeWithoutCapture()).readFile('sheet:Q3');
    expect(v).toBeNull();
  });

  it('returns [] from search when searchDocument is absent', async () => {
    const hits = await docMount(bridgeWithoutCapture()).search('', 'anything');
    expect(hits).toEqual([]);
  });

  it('truncates by BYTES, not characters, for multi-byte outline text', async () => {
    // Each '€' is 3 UTF-8 bytes; a naive text.slice(0, maxBytes) would keep far more than
    // maxBytes bytes here, violating the ReadOpts contract.
    const bridge = {
      surface: 'excel',
      async captureDocState() {
        return {
          surface: 'excel',
          version: 1,
          capturedAt: new Date(0).toISOString(),
          outline: [{ level: 1, text: '€€€€€€' }],
          inventory: [],
        };
      },
    } as unknown as DocBridge;
    const v = await docMount(bridge).readFile('outline.md', { maxBytes: 7 });
    expect(v?.truncated).toBe(true);
    expect(v?.bytes).toBeLessThanOrEqual(7);
    expect(new TextEncoder().encode(v!.text).length).toBe(v?.bytes);
  });
});
