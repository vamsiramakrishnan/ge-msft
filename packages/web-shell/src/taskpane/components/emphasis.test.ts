import { describe, expect, it } from 'vitest';
import { parseEmphasis, type EmphasisNode } from './emphasis.js';

/** Render the parse as compact HTML-ish text for readable expectations. */
function show(text: string): string {
  const render = (nodes: EmphasisNode<string>[]): string =>
    nodes
      .map((n) =>
        n.kind === 'text'
          ? n.text
          : n.kind === 'atom'
            ? `[${n.atom}]`
            : `<${n.kind}>${render(n.children)}</${n.kind}>`,
      )
      .join('');
  return render(parseEmphasis<string>([{ kind: 'text', text }]));
}

describe('parseEmphasis (CommonMark delimiter rules)', () => {
  it('renders plain em and strong', () => {
    expect(show('a *b* c')).toBe('a <em>b</em> c');
    expect(show('a **b** c')).toBe('a <strong>b</strong> c');
    expect(show('_b_ and __c__')).toBe('<em>b</em> and <strong>c</strong>');
  });

  it('renders triple markers as strong and em, no stray asterisks', () => {
    expect(show('***Gemini Enterprise > Data stores***.')).toBe(
      '<em><strong>Gemini Enterprise > Data stores</strong></em>.',
    );
  });

  it('renders em nested inside em (real Outlook answer, 2026-10-17)', () => {
    expect(show('Pilih *Microsoft Outlook, verifikasi masuk (*login*), aktif (*Active*)*.')).toBe(
      'Pilih <em>Microsoft Outlook, verifikasi masuk (<em>login</em>), aktif (<em>Active</em>)</em>.',
    );
  });

  it('keeps unmatched or non-flanking markers as text', () => {
    expect(show('Saya *ContosoAI, asisten Anda')).toBe('Saya *ContosoAI, asisten Anda');
    expect(show('5 * 3 * 2')).toBe('5 * 3 * 2');
    expect(show('Mail_Read and snake_case_name')).toBe('Mail_Read and snake_case_name');
  });

  it('treats atoms (code, links) as words next to a marker', () => {
    const nodes = parseEmphasis<string>([
      { kind: 'text', text: 'see *' },
      { kind: 'atom', atom: 'code' },
      { kind: 'text', text: '* now' },
    ]);
    expect(nodes).toEqual([
      { kind: 'text', text: 'see ' },
      { kind: 'em', children: [{ kind: 'atom', atom: 'code' }] },
      { kind: 'text', text: ' now' },
    ]);
  });
});
