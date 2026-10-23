// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { ConversationHistoryPanel } from './ConversationHistoryPanel.js';
import type { ConversationItem } from '../../controller.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

const chat = (id: string, surface?: string): ConversationItem => ({
  name: `sessions/${id}`,
  id,
  title: `chat ${id}`,
  turnCount: 1,
  isPinned: false,
  active: false,
  ...(surface ? { surface } : {}),
});

let resumed: string[] = [];

function render(items: ConversationItem[]): void {
  resumed = [];
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(
      createElement(ConversationHistoryPanel, {
        conversations: { items, loading: false, loaded: true },
        surface: 'excel',
        onRefresh: () => undefined,
        onResume: (name: string) => void resumed.push(name),
      }),
    );
  });
}

const titles = (): string[] =>
  [...container.querySelectorAll('.session-title')].map((n) => n.textContent ?? '');

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

describe('ConversationHistoryPanel — this app only', () => {
  it("shows only this app's chats by default, each tagged with its app", () => {
    render([chat('a', 'excel'), chat('b', 'outlook'), chat('c')]);
    expect(titles()).toEqual(['chat a']);
    expect(container.querySelector('.session-sub')?.textContent).toMatch(/^Excel · 1 turn/);
    const toggle = container.querySelector<HTMLInputElement>('.sessions-filter input')!;
    expect(toggle.checked).toBe(true);

    act(() => toggle.click());
    expect(titles()).toEqual(['chat a', 'chat b', 'chat c']);
    expect(container.querySelectorAll('.session-sub')[1]?.textContent).toMatch(/^Outlook · /);
  });

  it('says where the other chats are when this app has none', () => {
    render([chat('b', 'outlook'), chat('c', 'word')]);
    expect(titles()).toEqual([]);
    expect(container.querySelector('.sessions-empty')?.textContent).toMatch(
      /No Excel chats yet\. Untick "Excel only" to see 2 from other apps\./,
    );
  });

  it('asks before continuing a chat from another app; cancel keeps the current chat', () => {
    render([chat('a', 'excel'), chat('b', 'outlook')]);
    act(() => container.querySelector<HTMLInputElement>('.sessions-filter input')!.click());
    const continueButtons = () => [
      ...container.querySelectorAll<HTMLButtonElement>('.session-row .session-resume'),
    ];

    act(() => continueButtons()[0]!.click()); // same app: no question
    expect(resumed).toEqual(['sessions/a']);

    act(() => continueButtons()[1]!.click());
    const dialog = container.querySelector('.sessions-confirm[role="alertdialog"]');
    expect(dialog?.textContent).toContain(
      'This chat started in Outlook. Its answers are about that email. Continue it here in Excel?',
    );
    act(() =>
      [...dialog!.querySelectorAll('button')].find((b) => b.textContent === 'cancel')!.click(),
    );
    expect(resumed).toEqual(['sessions/a']);
    expect(container.querySelector('.sessions-confirm')).toBeNull();

    act(() => continueButtons()[1]!.click());
    act(() =>
      [...container.querySelectorAll('.sessions-confirm button')]
        .find((b) => b.textContent === 'continue here')!
        .dispatchEvent(new MouseEvent('click', { bubbles: true })),
    );
    expect(resumed).toEqual(['sessions/a', 'sessions/b']);
  });
});
