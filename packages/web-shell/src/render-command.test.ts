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

  it('uses placeholders for a write-cells with no target/value', () => {
    expect(renderCommandLine(req('write-cells', {}))).toBe('set <range>');
  });
});
