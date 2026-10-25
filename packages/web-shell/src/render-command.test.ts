import { describe, it, expect } from 'vitest';
import type { ActuationRequest } from '@ge/contracts';
import { asChangeId } from '@ge/contracts';
import { renderCommandLine } from './render-command.js';

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

  it('does not throw on a kind without a dedicated renderer; degrades to a best-effort label', () => {
    const out = renderCommandLine(req('insert-text', { text: 'hello there' }));
    expect(out).toBe('insert-text "hello there"');
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
