import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from 'react';
import type { Intent, QuickAction, Surface } from '@ge/contracts';
import type { ContextChip, ConversationsState, Skill } from '../../controller.js';
import { ContextTray } from './ContextTray.js';
import { SkillsPanel } from './SkillsPanel.js';
import { ActionLibrary } from './ActionLibrary.js';
import { ConversationHistoryPanel } from './ConversationHistoryPanel.js';

export interface ToolbarProps {
  surface: Surface;
  allowedIntents?: Iterable<Intent>;
  agentLabel?: string;
  /** Combined busy/gate block for actions (state.busy || hasBlockingGate). */
  busy: boolean;
  hasGate: boolean;
  chips: ContextChip[];
  attachedCount: number;
  availableCount: number;
  messageCount: number;
  proposalCount: number;
  skills: Skill[];
  conversations: ConversationsState;
  /** Ids already shown as primary actions, excluded from the quick-action catalog (no dupes). */
  primaryActionIds: string[];
  hasSettings: boolean;
  settingsPanel?: ReactNode;
  onToggleChip: (id: string, attach: boolean) => void;
  onRevealChip: (id: string) => void;
  onRefreshContext: () => void;
  onRefreshConversations: () => void;
  onResumeConversation: (name: string) => void;
  onInvokeSkill: (name: string, args: Record<string, string>) => void;
  onQuickAction: (action: QuickAction) => void;
}

type Panel = 'context' | 'actions' | 'skills' | 'sessions' | 'settings';

const HOST_GLYPH: Readonly<Record<Surface, string>> = {
  word: 'W',
  excel: 'X',
  powerpoint: 'P',
  onenote: 'N',
  outlook: 'O',
  teams: 'T',
};

const HOST_NAME: Readonly<Record<Surface, string>> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
  onenote: 'OneNote',
  outlook: 'Outlook',
  teams: 'Teams',
};

const PANEL_TITLE: Readonly<Record<Panel, string>> = {
  context: 'Context and grounding',
  actions: 'Find your next action',
  skills: 'Session skills',
  sessions: 'Conversation history',
  settings: 'Catalog and routing',
};

/**
 * The single icon toolbar that replaces the stacked chrome: product mark, host glyph + readiness
 * dot, and one icon per disclosure (Context / Actions / Skills) plus a settings gear. Each icon
 * opens one in-pane modal holding the existing section component in `embedded` mode. The modal owns
 * the available viewport and its own scrolling, so Office's narrow task panes cannot clip long quick
 * task or routing lists behind the composer. Closing: backdrop, close button, or Escape.
 */
export function Toolbar({
  surface,
  allowedIntents,
  agentLabel,
  busy,
  hasGate,
  chips,
  attachedCount,
  availableCount,

  skills,
  conversations,
  hasSettings,
  settingsPanel,
  onToggleChip,
  onRevealChip,
  onRefreshContext,
  onRefreshConversations,
  onResumeConversation,
  onInvokeSkill,
  onQuickAction,
}: ToolbarProps): JSX.Element {
  const [panel, setPanel] = useState<Panel | null>(null);
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

  useEffect(() => {
    if (!panel) return undefined;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    const focusTimer = window.setTimeout(() => {
      dialogRef.current?.querySelector<HTMLButtonElement>('.tw-modal-close')?.focus();
    }, 0);
    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') close();
    };
    document.addEventListener('keydown', onKeyDown, true);
    return () => {
      window.clearTimeout(focusTimer);
      document.body.style.overflow = previousOverflow;
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
        const trigger = rootRef.current?.querySelector<HTMLButtonElement>('[aria-label="Actions"]');
        if (trigger) choose('actions', trigger);
      }
    };
    document.addEventListener('keydown', shortcut);
    return () => document.removeEventListener('keydown', shortcut);
  }, [choose]);

  const state = hasGate ? 'gate' : busy ? 'busy' : 'ready';
  const status = hasGate ? 'Decision needed' : busy ? 'Working' : 'Ready';
  const hasSkills = skills.length > 0;

  return (
    <div className="tw" ref={rootRef}>
      <div className="tw-identity">
        <span className="tw-brand" aria-hidden="true" />
        <span className="tw-name" title={agentLabel ?? 'Grounded on your research unit'}>
          Gemini Enterprise
        </span>
        <span className="tw-spacer" />
        <span
          className="tw-host"
          title={`${HOST_NAME[surface]} · ${status}`}
          aria-label={`${HOST_NAME[surface]}, ${status}`}
        >
          <span className="tw-host-glyph" aria-hidden="true">
            {HOST_GLYPH[surface]}
          </span>
          <span className="tw-host-dot" data-state={state} aria-hidden="true" />
          <span className="tw-host-name">{HOST_NAME[surface]}</span>
        </span>
      </div>

      <div className="tw-bar" role="toolbar" aria-label="Gemini Enterprise controls">
        <button
          type="button"
          className={`tw-icon${panel === 'context' ? ' on' : ''}`}
          aria-expanded={panel === 'context'}
          aria-haspopup="dialog"
          aria-controls="tw-panel-context"
          aria-label={`Context — ${attachedCount} attached, ${availableCount} nearby`}
          title="Context & grounding"
          onClick={(event) => choose('context', event.currentTarget)}
        >
          <ToolbarIcon name="context" />
          <span className="tw-label">Context</span>
          {attachedCount > 0 && <span className="tw-badge">{attachedCount}</span>}
        </button>

        <button
          type="button"
          className={`tw-icon${panel === 'actions' ? ' on' : ''}`}
          aria-expanded={panel === 'actions'}
          aria-haspopup="dialog"
          aria-controls="tw-panel-actions"
          aria-label="Actions"
          title="Find actions (Ctrl / ⌘ K)"
          onClick={(event) => choose('actions', event.currentTarget)}
        >
          <ToolbarIcon name="actions" />
          <span className="tw-label">Actions</span>
        </button>

        {hasSkills && (
          <button
            type="button"
            className={`tw-icon${panel === 'skills' ? ' on' : ''}`}
            aria-expanded={panel === 'skills'}
            aria-haspopup="dialog"
            aria-controls="tw-panel-skills"
            aria-label={`Skills — ${skills.length} registered`}
            title="Session skills"
            onClick={(event) => choose('skills', event.currentTarget)}
          >
            <ToolbarIcon name="skills" />
            <span className="tw-label">Skills</span>
            <span className="tw-badge">{skills.length}</span>
          </button>
        )}

        <button
          type="button"
          className={`tw-icon${panel === 'sessions' ? ' on' : ''}`}
          aria-expanded={panel === 'sessions'}
          aria-haspopup="dialog"
          aria-controls="tw-panel-sessions"
          aria-label={`Conversations — ${conversations.items.length} loaded`}
          title="Conversations"
          onClick={(event) => choose('sessions', event.currentTarget)}
        >
          <ToolbarIcon name="sessions" />
          <span className="tw-label">History</span>
          {conversations.items.length > 0 && (
            <span className="tw-badge">{conversations.items.length}</span>
          )}
        </button>

        {hasSettings && (
          <button
            type="button"
            className={`tw-icon${panel === 'settings' ? ' on' : ''}`}
            aria-expanded={panel === 'settings'}
            aria-haspopup="dialog"
            aria-controls="tw-panel-settings"
            aria-label="Catalog and routing settings"
            title="Routing"
            onClick={(event) => choose('settings', event.currentTarget)}
          >
            <ToolbarIcon name="settings" />
            <span className="tw-label">Routing</span>
          </button>
        )}
      </div>

      {/* Keep every pane mounted so context/catalog state remains warm between disclosures. */}
      <div
        className="tw-modal-layer"
        hidden={!panel}
        onPointerDown={(event) => {
          if (event.target === event.currentTarget) close();
        }}
      >
        <section
          id="tw-modal-dialog"
          ref={dialogRef}
          className="tw-modal"
          role="dialog"
          aria-modal="true"
          aria-labelledby="tw-modal-title"
          onKeyDown={trapFocus}
        >
          <header className="tw-modal-head">
            <div>
              <span className="eyebrow">Gemini Enterprise</span>
              <h2 id="tw-modal-title" className="tw-modal-title">
                {panel ? PANEL_TITLE[panel] : 'Commands'}
              </h2>
            </div>
            <button type="button" className="tw-modal-close" aria-label="Close" onClick={close}>
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path d="M6 6l12 12M18 6 6 18" />
              </svg>
            </button>
          </header>

          <div className="tw-modal-body">
            <div id="tw-panel-context" className="tw-modal-pane" hidden={panel !== 'context'}>
              <ContextTray
                embedded
                disabled={busy}
                chips={chips}
                onToggle={onToggleChip}
                onReveal={onRevealChip}
                onRefresh={onRefreshContext}
              />
            </div>

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
            </div>

            <div id="tw-panel-skills" className="tw-modal-pane" hidden={panel !== 'skills'}>
              <SkillsPanel
                embedded
                skills={skills}
                disabled={busy}
                onInvoke={(name, args) => {
                  onInvokeSkill(name, args);
                  close();
                }}
              />
            </div>

            <div id="tw-panel-sessions" className="tw-modal-pane" hidden={panel !== 'sessions'}>
              <ConversationHistoryPanel
                conversations={conversations}
                disabled={busy}
                onRefresh={onRefreshConversations}
                onResume={(name) => {
                  onResumeConversation(name);
                  close();
                }}
              />
            </div>

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

type ToolbarIconName = 'context' | 'actions' | 'skills' | 'sessions' | 'settings';

function ToolbarIcon({ name }: { name: ToolbarIconName }): JSX.Element {
  const paths: Record<ToolbarIconName, JSX.Element> = {
    context: (
      <>
        <circle cx="12" cy="12" r="3" />
        <circle cx="5" cy="6" r="2" />
        <circle cx="19" cy="6" r="2" />
        <path d="M6.7 7.2 9.4 10M17.3 7.2 14.6 10M12 15v4" />
      </>
    ),
    actions: <path d="m13 2-8 12h6l-1 8 9-13h-6z" />,
    skills: (
      <>
        <path d="M4 5.5A2.5 2.5 0 0 1 6.5 3H11v16H6.5A2.5 2.5 0 0 0 4 21.5z" />
        <path d="M20 5.5A2.5 2.5 0 0 0 17.5 3H13v16h4.5a2.5 2.5 0 0 1 2.5 2.5z" />
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
