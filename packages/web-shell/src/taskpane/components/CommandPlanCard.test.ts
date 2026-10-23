// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { CommandPlan } from '@ge/contracts';
import { CommandPlanCard } from './CommandPlanCard.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

const OUTLOOK_ID =
  'AAMkAGY0NTg4ZGVhLWEwNmQtNDcyZC05MDMzLTgwYWE5M2ZhZDUzYwBGAAAAAACQx/XJfTPcQpSoXRrRytIaBwCj4mMuDwbRQLOrDNysSYShAAAAAAEJAAAArsHu7TpfAAA=';

function render(plan: Partial<CommandPlan>, language: 'en' | 'id' = 'en'): string {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  const full: CommandPlan = {
    intent: 'draft',
    surface: 'outlook',
    ground: [],
    context: [],
    steps: ['Stage a reply that accepts the meeting invitation politely'],
    excludes: [],
    clarify: [],
    ...plan,
  };
  act(() => {
    root.render(
      createElement(CommandPlanCard, {
        language,
        pending: { plan: full, task: 'reply' },
        onConfirm: vi.fn(),
        onCancel: vi.fn(),
      }),
    );
  });
  return container.textContent ?? '';
}

describe('CommandPlanCard sources', () => {
  it('names an email by what it is, never by its Outlook item id', () => {
    const text = render({ ground: [{ kind: 'document', ref: `mail:${OUTLOOK_ID}` }] });
    expect(text).not.toContain('AAMk');
    expect(text).toContain('Using an email');
  });

  it('says it in Bahasa when the chat is in Bahasa', () => {
    const text = render({ ground: [{ kind: 'document', ref: `mail:${OUTLOOK_ID}` }] }, 'id');
    expect(text).toContain('Menggunakan sebuah email');
  });

  it('names any other long internal id as "an item", in sources and in the scope', () => {
    const text = render({
      ground: [{ kind: 'document', ref: 'shape:{9F2B6C1E-44A0-4F3B-9C7D-2E5A8B1C0D3F}' }],
      scope: { kind: 'section', ref: `item:${OUTLOOK_ID}` },
    });
    expect(text).not.toContain('9F2B6C1E');
    expect(text).not.toContain('AAMk');
    expect(text).toContain('In an item');
    expect(text).toContain('Using an item');
  });

  it('keeps readable sources: sheet names, cell ranges and named documents', () => {
    const text = render({
      ground: [
        { kind: 'document', ref: 'Sales!A1:D7' },
        { kind: 'document', ref: 'Vendor Risk notebook' },
      ],
    });
    expect(text).toContain('Sales!A1:D7');
    expect(text).toContain('Vendor Risk notebook');
  });
});
