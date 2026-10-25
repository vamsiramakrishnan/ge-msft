import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from 'react';
import type { Intent, QuickAction, Surface } from '@ge/contracts';
import type { ConversationsState } from '../../controller.js';
import { ActionLibrary } from './ActionLibrary.js';
import { ConversationHistoryPanel } from './ConversationHistoryPanel.js';
import { brand } from '../../brand.js';
import type { SavedDiagnostic } from '../../saved-diagnostics.js';
import { SavedDiagnosticsPanel } from './SavedDiagnosticsPanel.js';

export interface ToolbarProps {
  surface: Surface;
  allowedIntents?: Iterable<Intent>;
  agentLabel?: string;
  /** Combined busy/gate block for actions (state.busy || hasBlockingGate). */
  busy: boolean;
  hasGate: boolean;
  messageCount: number;
  proposalCount: number;
  conversations: ConversationsState;
  /** Ids already shown as primary actions, excluded from the quick-action catalog (no dupes). */
  primaryActionIds: string[];
  hasSettings: boolean;
  settingsPanel?: ReactNode;
  /** Extra tools shown under the actions in the Tools panel (e.g. Excel workflows). */
  tools?: ReactNode;
  /** The pane's main column, shown beside the sidebar. */
  children?: ReactNode;
  onRefreshConversations: () => void;
  onResumeConversation: (name: string) => void;
  onQuickAction: (action: QuickAction) => void;
  /** Start a new chat in place; the button is hidden when not provided. */
  onNewConversation?: () => void;
  /** Earlier requests' diagnostics; the ⋯ button shows only when there are some. */
  savedDiagnostics?: SavedDiagnostic[];
  onClearSavedDiagnostics?: () => void;
}

type Panel = 'actions' | 'sessions' | 'settings' | 'diagnostics';

const PANEL_TITLE: Readonly<Record<Panel, string>> = {
  actions: 'Tools',
  sessions: 'History',
  settings: 'Settings',
  diagnostics: 'Earlier diagnostics',
};

/** The sidebar's open/closed choice is a per-viewer convenience; storage may be unavailable. */
const RAIL_KEY = 'ge.sidebar';
/** Below this pane width the sidebar starts collapsed, until the user opens it themselves. */
const NARROW_PANE = 360;
function readRailOpen(): boolean {
  let saved: string | null = null;
  try {
    saved = window.localStorage.getItem(RAIL_KEY);
  } catch {
    // Storage blocked: fall back to the pane width.
  }
  if (saved) return saved !== 'closed';
  return window.innerWidth >= NARROW_PANE;
}
function writeRailOpen(open: boolean): void {
  try {
    window.localStorage.setItem(RAIL_KEY, open ? 'open' : 'closed');
  } catch {
    // Storage blocked: the choice lasts until the pane closes.
  }
}

/**
 * The pane shell: a collapsible left sidebar (Chat / Tools / History / Settings) beside the
 * main column, which starts with a slim header (mark, name, new chat, ⋯). Each sidebar item opens one
 * panel over the main column holding the existing section component; Chat closes it. Closing: the
 * close button, Escape, or Chat.
 */
export function Toolbar({
  surface,
  allowedIntents,
  busy,
  hasGate,
  conversations,
  hasSettings,
  settingsPanel,
  tools,
  children,
  onRefreshConversations,
  onResumeConversation,
  onQuickAction,
  onNewConversation,
  savedDiagnostics = [],
  onClearSavedDiagnostics,
}: ToolbarProps): JSX.Element {
  const [panel, setPanel] = useState<Panel | null>(null);
  const [railOpen, setRailOpen] = useState(readRailOpen);
  const rootRef = useRef<HTMLDivElement>(null);
  const dialogRef = useRef<HTMLElement>(null);
  const lastTriggerRef = useRef<HTMLButtonElement | null>(null);
  const close = useCallback(() => {
    setPanel(null);
    window.setTimeout(() => lastTriggerRef.current?.focus(), 0);
  }, []);
  const choose = useCallback(
    (next: Panel, trigger: HTMLButtonElement) => {
      if (panel === next) {
        close();
        return;
      }
      lastTriggerRef.current = trigger;
      setPanel(next);
    },
    [close, panel],
  );
  const toggleRail = (): void => {
    setRailOpen((open) => {
      writeRailOpen(!open);
      return !open;
    });
  };

  // A decision card (plan, write or share approval) sits in the main column: an open panel would
  // cover it, e.g. when a workflow run from Tools stages its write. Close the panel so it is seen.
  useEffect(() => {
    if (hasGate) setPanel(null);
  }, [hasGate]);

  useEffect(() => {
    if (!panel) return undefined;
    const focusTimer = window.setTimeout(() => {
      dialogRef.current?.querySelector<HTMLButtonElement>('.tw-modal-close')?.focus();
    }, 0);
    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') close();
    };
    document.addEventListener('keydown', onKeyDown, true);
    return () => {
      window.clearTimeout(focusTimer);
      document.removeEventListener('keydown', onKeyDown, true);
    };
  }, [close, panel]);

  const trapFocus = useCallback((event: ReactKeyboardEvent<HTMLElement>): void => {
    if (event.key !== 'Tab') return;
    const focusable = [
      ...(dialogRef.current?.querySelectorAll<HTMLElement>(
        'button:not(:disabled), a[href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
      ) ?? []),
    ].filter((node) => !node.closest('[hidden]'));
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (!first || !last) return;
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  }, []);

  useEffect(() => {
    if (panel === 'sessions' && !conversations.loaded && !conversations.loading) {
      onRefreshConversations();
    }
  }, [conversations.loaded, conversations.loading, onRefreshConversations, panel]);

  useEffect(() => {
    const shortcut = (event: KeyboardEvent): void => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') {
        event.preventDefault();
        const trigger = rootRef.current?.querySelector<HTMLButtonElement>('[data-panel="actions"]');
        if (trigger) choose('actions', trigger);
      }
    };
    document.addEventListener('keydown', shortcut);
    return () => document.removeEventListener('keydown', shortcut);
  }, [choose]);

  const railItem = (
    name: Panel,
    icon: ToolbarIconName,
    label: string,
    title: string,
    badge?: number,
  ): JSX.Element => (
    <button
      type="button"
      className={`tw-rail-item${panel === name ? ' on' : ''}`}
      data-panel={name}
      aria-expanded={panel === name}
      aria-controls={`tw-panel-${name}`}
      title={title}
      onClick={(event) => choose(name, event.currentTarget)}
    >
      <ToolbarIcon name={icon} />
      <span className="tw-label">{label}</span>
      {badge ? <span className="tw-badge">{badge}</span> : null}
    </button>
  );

  return (
    <div className={`tw-shell${railOpen ? '' : ' rail-closed'}`} ref={rootRef}>
      {railOpen && (
        <nav className="tw-rail" aria-label={`${brand.name} sidebar`}>
          <button
            type="button"
            className="tw-rail-toggle"
            aria-label="Collapse sidebar"
            title="Collapse sidebar"
            onClick={toggleRail}
          >
            <ToolbarIcon name="collapse" />
          </button>
          <button
            type="button"
            className={`tw-rail-item${panel === null ? ' on' : ''}`}
            aria-current={panel === null ? 'page' : undefined}
            title="Chat"
            onClick={() => panel && close()}
          >
            <ToolbarIcon name="chat" />
            <span className="tw-label">Chat</span>
          </button>
          {railItem('actions', 'tools', 'Tools', 'Tools (Ctrl / ⌘ K)')}
          {railItem('sessions', 'sessions', 'History', 'Conversation history')}
          <span className="tw-rail-spacer" />
          {hasSettings && railItem('settings', 'settings', 'Settings', 'Answer format and style')}
        </nav>
      )}

      <div className="tw-main">
        <header className="tw-head" role="toolbar" aria-label={`${brand.name} controls`}>
          {!railOpen && (
            <button
              type="button"
              className="tw-head-btn"
              aria-label="Open sidebar"
              title="Open sidebar"
              onClick={toggleRail}
            >
              <ToolbarIcon name="menu" />
            </button>
          )}
          <span className="tw-brand" role="img" aria-label={brand.name} title={brand.name} />
          <span className="tw-name">{brand.name}</span>
          {onNewConversation && (
            <button
              type="button"
              className="tw-head-btn"
              aria-label="New chat"
              title="Start a new chat"
              disabled={busy}
              onClick={onNewConversation}
            >
              <ToolbarIcon name="new" />
            </button>
          )}
          {savedDiagnostics.length > 0 && (
            <button
              type="button"
              className={`tw-head-btn${panel === 'diagnostics' ? ' on' : ''}`}
              aria-expanded={panel === 'diagnostics'}
              aria-controls="tw-panel-diagnostics"
              aria-label="Earlier diagnostics"
              title="Earlier diagnostics"
              onClick={(event) => choose('diagnostics', event.currentTarget)}
            >
              <ToolbarIcon name="more" />
            </button>
          )}
        </header>

        {children}

        {/* Keep every pane mounted so catalog and history state remain warm between disclosures. */}
        <section
          id="tw-modal-dialog"
          ref={dialogRef}
          className="tw-modal"
          hidden={!panel}
          aria-labelledby="tw-modal-title"
          onKeyDown={trapFocus}
        >
          <header className="tw-modal-head">
            <h2 id="tw-modal-title" className="tw-modal-title">
              {panel ? PANEL_TITLE[panel] : ''}
            </h2>
            <button type="button" className="tw-modal-close" aria-label="Close" onClick={close}>
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path d="M6 6l12 12M18 6 6 18" />
              </svg>
            </button>
          </header>

          <div className="tw-modal-body">
            <div id="tw-panel-actions" className="tw-modal-pane" hidden={panel !== 'actions'}>
              <ActionLibrary
                surface={surface}
                allowedIntents={allowedIntents}
                disabled={busy}
                onAction={(action) => {
                  onQuickAction(action);
                  close();
                }}
              />
              {tools}
            </div>

            <div id="tw-panel-sessions" className="tw-modal-pane" hidden={panel !== 'sessions'}>
              <ConversationHistoryPanel
                conversations={conversations}
                surface={surface}
                disabled={busy}
                onRefresh={onRefreshConversations}
                onResume={(name) => {
                  onResumeConversation(name);
                  close();
                }}
              />
            </div>

            {savedDiagnostics.length > 0 ? (
              <div
                id="tw-panel-diagnostics"
                className="tw-modal-pane"
                hidden={panel !== 'diagnostics'}
              >
                <SavedDiagnosticsPanel
                  entries={savedDiagnostics}
                  {...(onClearSavedDiagnostics
                    ? {
                        onClear: () => {
                          onClearSavedDiagnostics();
                          close();
                        },
                      }
                    : {})}
                />
              </div>
            ) : null}

            {hasSettings ? (
              <div id="tw-panel-settings" className="tw-modal-pane" hidden={panel !== 'settings'}>
                {settingsPanel}
              </div>
            ) : null}
          </div>
        </section>
      </div>
    </div>
  );
}

type ToolbarIconName =
  | 'new'
  | 'chat'
  | 'tools'
  | 'sessions'
  | 'settings'
  | 'collapse'
  | 'menu'
  | 'more';

function ToolbarIcon({ name }: { name: ToolbarIconName }): JSX.Element {
  const paths: Record<ToolbarIconName, JSX.Element> = {
    new: (
      <>
        <path d="M12 4H5a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-7" />
        <path d="M17 3l4 4-9 9H8v-4z" />
      </>
    ),
    chat: (
      <path d="M4 5.5A1.5 1.5 0 0 1 5.5 4h13A1.5 1.5 0 0 1 20 5.5v9a1.5 1.5 0 0 1-1.5 1.5H10l-6 4z" />
    ),
    tools: (
      <>
        <rect x="4" y="4" width="6.5" height="6.5" rx="1" />
        <rect x="13.5" y="4" width="6.5" height="6.5" rx="1" />
        <rect x="4" y="13.5" width="6.5" height="6.5" rx="1" />
        <rect x="13.5" y="13.5" width="6.5" height="6.5" rx="1" />
      </>
    ),
    collapse: (
      <>
        <rect x="3" y="4" width="18" height="16" rx="2" />
        <path d="M9 4v16M15.5 10 13.5 12l2 2" />
      </>
    ),
    menu: <path d="M4 7h16M4 12h16M4 17h16" />,
    more: (
      <>
        <circle cx="5" cy="12" r="1" />
        <circle cx="12" cy="12" r="1" />
        <circle cx="19" cy="12" r="1" />
      </>
    ),
    sessions: (
      <>
        <circle cx="12" cy="12" r="8.5" />
        <path d="M12 7.5V12l3 2" />
      </>
    ),
    settings: (
      <>
        <path d="M4 7h10M18 7h2M4 17h2M10 17h10" />
        <circle cx="16" cy="7" r="2" />
        <circle cx="8" cy="17" r="2" />
      </>
    ),
  };
  return (
    <svg className="tw-glyph" viewBox="0 0 24 24" aria-hidden="true">
      {paths[name]}
    </svg>
  );
}
