import { describe, it, expect } from 'vitest';
import type { ActuationRequest } from '@ge/contracts';
import { asChangeId, parseCommandLine } from '@ge/contracts';
import { compileCommand } from '@ge/runtime';
import { renderCommandLine, textActuationTargetLabel } from './render-command.js';

const req = (
  kind: ActuationRequest['kind'],
  params: ActuationRequest['params'],
): ActuationRequest => ({
  changeId: asChangeId('c-1'),
  kind,
  surface: 'word',
  params,
});

describe('renderCommandLine', () => {
  it('shows a reply file and the full text it holds (security review 2026-10-17)', () => {
    expect(
      renderCommandLine(
        req('reply-mail', {
          mail: { body: 'See attached.' },
          attachment: { name: 'notes.docx', base64: 'UEsDBA==', text: 'Q3 "draft" figures' },
        }),
      ),
    ).toBe('mail "See attached." attach="notes.docx" text="Q3 \\"draft\\" figures"');
  });

  it('renders write-cells as `set <cell> <value>` (formula first)', () => {
    expect(
      renderCommandLine(req('write-cells', { target: { range: 'Sales!F2' }, cells: [['=C2-D2']] })),
    ).toBe('set Sales!F2 =C2-D2');
  });

  it('renders write-cells with a literal value', () => {
    expect(
      renderCommandLine(req('write-cells', { target: { range: 'Sales!A1' }, cells: [['42']] })),
    ).toBe('set Sales!A1 42');
  });

  it('renders multi-cell write-cells as a grid summary', () => {
    expect(
      renderCommandLine(
        req('write-cells', {
          target: { range: "'Daily schedule'!C5:D6" },
          cells: [
            ['Monday', 'Tuesday'],
            ['Deep Work', 'Music Lesson'],
          ],
        }),
      ),
    ).toBe("grid 'Daily schedule'!C5:D6 (2x2)");
  });

  it('renders tracked-change as `suggest "old" => "new"`', () => {
    expect(
      renderCommandLine(req('tracked-change', { target: { matchText: 'old' }, text: 'new' })),
    ).toBe('suggest "old" => "new"');
  });

  it('escapes quotes and backslashes in suggest operands', () => {
    expect(
      renderCommandLine(req('tracked-change', { target: { matchText: 'a "b"' }, text: 'c\\d' })),
    ).toBe('suggest "a \\"b\\"" => "c\\\\d"');
  });

  it('renders add-comment as `comment <sel> "text"`', () => {
    expect(
      renderCommandLine(
        req('add-comment', { target: { matchText: 'clause 3' }, text: 'check this' }),
      ),
    ).toBe('comment clause 3 "check this"');
  });

  it('renders comment-reply best-effort', () => {
    expect(
      renderCommandLine(req('comment-reply', { target: { commentId: 'cmt-9' }, text: 'agreed' })),
    ).toBe('comment cmt-9 "agreed"');
  });

  it('renders format-cells as `format <range> k=v ...`', () => {
    expect(
      renderCommandLine(
        req('format-cells', {
          target: { range: 'A1:A3' },
          format: { bold: true, fill: '#FFF2CC' },
        }),
      ),
    ).toBe('format A1:A3 bold=true fill=#FFF2CC');
  });

  it('discloses the native selection replacement default for an unanchored text insert', () => {
    const out = renderCommandLine(req('insert-text', { text: 'hello there' }));
    expect(out).toBe('/insert-text position=replace text="hello there"');
  });

  it('degrades to the bare kind when no hint is present', () => {
    expect(renderCommandLine(req('insert-slide', {}))).toBe('insert-slide');
  });

  it('renders insert-slide as the slide command, with at= 1-based', () => {
    expect(
      renderCommandLine(
        req('insert-slide', {
          slide: { title: 'Roadmap', bullets: ['Q4 2026: pilot'] },
          target: { slideIndex: 2 },
        }),
      ),
    ).toBe('slide "Roadmap" "Q4 2026: pilot" at=3');
  });

  it('never labels a deck insert as a single slide, even with slide params', () => {
    const line = renderCommandLine(
      req('insert-slide', { slide: { title: 'Decoy', bullets: [] }, ooxml: 'UEsDB...' }),
    );
    expect(line).not.toContain('slide "Decoy"');
  });

  it('uses placeholders for a write-cells with no target/value', () => {
    expect(renderCommandLine(req('write-cells', {}))).toBe('set <range>');
  });
});

describe('native text approvals preserve the actual scope and pins', () => {
  const pin = 'sha256:' + 'a'.repeat(64);
  const value = '"புதிய உரை"\t\\\n😀\u2028```cmd\n/insert-text text="other"\n```\u202e';
  it.each([
    {
      kind: 'insert-text' as const,
      surface: 'word' as const,
      params: {
        position: 'end' as const,
        text: value,
        textFont: { bold: true, name: 'Aptos', size: 12 },
        target: { expectedDocumentId: pin },
      },
    },
    {
      kind: 'insert-text' as const,
      surface: 'word' as const,
      params: {
        position: 'start' as const,
        text: value,
        target: {
          matchText: 'Quoted "anchor"\nTamil தமிழ்',
          contextHint: 'the original clause',
          expectedDocumentId: pin,
        },
      },
    },
    {
      kind: 'fill-content-control' as const,
      surface: 'word' as const,
      params: {
        text: value,
        target: { contentControlId: '7', expectedDocumentId: pin, expectedTextHash: pin },
      },
    },
    {
      kind: 'set-shape-text' as const,
      surface: 'powerpoint' as const,
      params: {
        text: value,
        target: {
          slideId: 's1#7',
          shapeId: 'shape-3',
          expectedDocumentId: pin,
          expectedTextHash: pin,
        },
      },
    },
  ])(
    'round-trips $kind through the actual compiler without dropping targets, positions or Unicode',
    (input) => {
      const request = { ...input, changeId: asChangeId('rendered') };
      const rendered = renderCommandLine(request);
      expect(rendered.split('\n')).toHaveLength(1);
      expect(rendered).not.toContain('```');
      expect(rendered).not.toContain('\u202e');
      const parsed = parseCommandLine(rendered);
      if ('error' in parsed) throw new Error(parsed.error);
      const compiled = compileCommand(parsed, {
        surface: input.surface,
        mintChangeId: () => asChangeId('rendered'),
      });
      expect(compiled).toMatchObject({
        kind: 'write',
        request: { kind: input.kind, surface: input.surface, params: input.params },
      });
    },
  );
  it('distinguishes document boundaries, matched-range boundaries and selection replacement', () => {
    expect(textActuationTargetLabel(req('insert-text', { position: 'end', text: 'new' }))).toBe(
      'End of document body',
    );
    expect(
      textActuationTargetLabel(
        req('insert-text', {
          position: 'start',
          text: 'new',
          target: { matchText: 'anchor', contextHint: 'governing clause' },
        }),
      ),
    ).toBe('Start of matched text "anchor" · context "governing clause"');
    expect(textActuationTargetLabel(req('insert-text', { text: 'new' }))).toBe(
      'Replace current selection',
    );
    expect(
      textActuationTargetLabel(
        req('insert-text', { text: 'new', target: { matchText: 'anchor' } }),
      ),
    ).toBe('After matched text "anchor"');
  });
});

describe('the card shows what a sort or a table replacement does (security review H1, H3, L1)', () => {
  const req = (kind: ActuationRequest['kind'], params: ActuationRequest['params']) =>
    ({ changeId: asChangeId('c'), surface: 'excel', kind, params }) as ActuationRequest;

  it('renders the pinned sort range, key, order and headers', () => {
    expect(
      renderCommandLine(
        req('sort-range', {
          sortRange: {
            range: "'Branch Performance'!A1:F21",
            key: 'Deposits',
            order: 'descending',
            hasHeaders: true,
          },
        }),
      ),
    ).toBe(
      `/sort-range range='Branch Performance'!A1:F21 key="Deposits" order=descending headers=true`,
    );
  });

  it('says a table replaces paragraphs, and from which to which', () => {
    const line = renderCommandLine(
      req('insert-table', {
        target: { matchText: 'Budi - send' },
        tableGrid: {
          rows: [
            ['Owner', 'Action'],
            ['Budi', 'Send'],
          ],
          hasHeaders: true,
          replaceThrough: 'Rina - review',
        },
      }),
    );
    expect(line).toBe(
      '/insert-table 2x2 table (first row: Owner | Action) REPLACES the paragraphs from "Budi - send" through "Rina - review"',
    );
  });

  it('shows the outside and inside border weights', () => {
    expect(
      renderCommandLine(
        req('format-cells', {
          target: { range: 'A1:F50' },
          format: { borderOutside: 'thick', borderInside: 'none' },
        }),
      ),
    ).toBe('format A1:F50 borderOutside=thick borderInside=none');
  });
});
