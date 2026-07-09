import { useCallback, useEffect, useRef, useState } from 'react';
import type { Intent, QuickAction, Surface } from '@ge/contracts';
import type { ContextChip, ConversationsState, Skill } from '../../controller.js';
import { ContextTray } from './ContextTray.js';
import { SkillsPanel } from './SkillsPanel.js';
import { QuickActionBar } from './QuickActionBar.js';
import { SurfaceCommandCenter } from './SurfaceCommandCenter.js';
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
  onOpenSettings: () => void;
  onToggleChip: (id: string, attach: boolean) => void;
  onRevealChip: (id: string) => void;
  onRefreshContext: () => void;
  onRefreshConversations: () => void;
  onResumeConversation: (name: string) => void;
  onInvokeSkill: (name: string, args: Record<string, string>) => void;
  onQuickAction: (action: QuickAction) => void;
}

type Panel = 'context' | 'actions' | 'skills' | 'sessions';

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

/**
 * The single icon toolbar that replaces the stacked chrome: product mark, host glyph + readiness
 * dot, and one icon per disclosure (Context / Actions / Skills) plus a settings gear. Each icon
 * opens ONE fixed-position sheet (so it escapes the thread's `overflow:auto` clipping) holding the
 * existing section component in `embedded` mode — the sheet is the disclosure, the section renders
 * expanded. Settings opens the catalog modal owned by App. Closing: click-away or Escape. This is
 * pure chrome; the gate rail, thread, and composer stay where they are in App.
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
  messageCount,
  proposalCount,
  skills,
  conversations,
  primaryActionIds,
  hasSettings,
  onOpenSettings,
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
  const close = useCallback(() => setPanel(null), []);
  const choose = useCallback((next: Panel) => setPanel((cur) => (cur === next ? null : next)), []);

  useEffect(() => {
    if (!panel) return undefined;
    const onPointerDown = (e: PointerEvent): void => {
      const root = rootRef.current;
      if (root && !root.contains(e.target as Node)) setPanel(null);
    };
    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') setPanel(null);
    };
    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown, true);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown, true);
    };
  }, [panel]);

  useEffect(() => {
    if (panel === 'sessions' && !conversations.loaded && !conversations.loading) {
      onRefreshConversations();
    }
  }, [conversations.loaded, conversations.loading, onRefreshConversations, panel]);

  const state = hasGate ? 'gate' : busy ? 'busy' : 'ready';
  const status = hasGate ? 'Decision needed' : busy ? 'Working' : 'Ready';
  const hasSkills = skills.length > 0;

  return (
    <div className="tw" ref={rootRef}>
      <div className="tw-bar" role="toolbar" aria-label="Gemini Enterprise controls">
        <span className="tw-brand" aria-hidden="true" />
        <span className="tw-name" title={agentLabel ?? 'Grounded on your research unit'}>
          Gemini Enterprise
        </span>
        <span
          className="tw-host"
          title={`${HOST_NAME[surface]} · ${status}`}
          aria-label={`${HOST_NAME[surface]}, ${status}`}
        >
          <span className="tw-host-glyph" aria-hidden="true">
            {HOST_GLYPH[surface]}
          </span>
          <span className="tw-host-dot" data-state={state} aria-hidden="true" />
        </span>

        <span className="tw-spacer" />

        <button
          type="button"
          className={`tw-icon${panel === 'context' ? ' on' : ''}`}
          aria-expanded={panel === 'context'}
          aria-haspopup="dialog"
          aria-label={`Context — ${attachedCount} attached, ${availableCount} nearby`}
          title="Context & grounding"
          onClick={() => choose('context')}
        >
          <span className="tw-glyph" aria-hidden="true">
            ◎
          </span>
          {attachedCount > 0 && <span className="tw-badge">{attachedCount}</span>}
        </button>

        <button
          type="button"
          className={`tw-icon${panel === 'actions' ? ' on' : ''}`}
          aria-expanded={panel === 'actions'}
          aria-haspopup="dialog"
          aria-label="Actions"
          title="Actions"
          onClick={() => choose('actions')}
        >
          <span className="tw-glyph" aria-hidden="true">
            ⚡
          </span>
        </button>

        {hasSkills && (
          <button
            type="button"
            className={`tw-icon${panel === 'skills' ? ' on' : ''}`}
            aria-expanded={panel === 'skills'}
            aria-haspopup="dialog"
            aria-label={`Skills — ${skills.length} registered`}
            title="Session skills"
            onClick={() => choose('skills')}
          >
            <span className="tw-glyph" aria-hidden="true">
              ✦
            </span>
            <span className="tw-badge">{skills.length}</span>
          </button>
        )}

        <button
          type="button"
          className={`tw-icon${panel === 'sessions' ? ' on' : ''}`}
          aria-expanded={panel === 'sessions'}
          aria-haspopup="dialog"
          aria-label={`Conversations — ${conversations.items.length} loaded`}
          title="Conversations"
          onClick={() => choose('sessions')}
        >
          <span className="tw-glyph" aria-hidden="true">
            ◷
          </span>
          {conversations.items.length > 0 && (
            <span className="tw-badge">{conversations.items.length}</span>
          )}
        </button>

        {hasSettings && (
          <button
            type="button"
            className="tw-icon"
            aria-haspopup="dialog"
            aria-label="Catalog and routing settings"
            title="Settings"
            onClick={onOpenSettings}
          >
            <span className="tw-glyph" aria-hidden="true">
              ⚙
            </span>
          </button>
        )}
      </div>

      {/* Sheets stay mounted (CSS-hidden when inactive) so the section DOM is always present —
          only one is visible at a time via data-active. */}
      <div
        className="tw-sheet"
        role="dialog"
        aria-label="Context and grounding"
        data-active={panel === 'context' ? 'true' : 'false'}
        hidden={panel !== 'context'}
      >
        <ContextTray
          embedded
          chips={chips}
          onToggle={onToggleChip}
          onReveal={onRevealChip}
          onRefresh={onRefreshContext}
        />
      </div>

      <div
        className="tw-sheet"
        role="dialog"
        aria-label="Actions"
        data-active={panel === 'actions' ? 'true' : 'false'}
        hidden={panel !== 'actions'}
      >
        <SurfaceCommandCenter
          surface={surface}
          allowedIntents={allowedIntents}
          busy={busy}
          hasGate={hasGate}
          attachedCount={attachedCount}
          availableCount={availableCount}
          messageCount={messageCount}
          proposalCount={proposalCount}
          onAction={(action) => {
            onQuickAction(action);
            close();
          }}
        />
        <QuickActionBar
          embedded
          surface={surface}
          allowedIntents={allowedIntents}
          busy={busy}
          excludeIds={primaryActionIds}
          onAction={(action) => {
            onQuickAction(action);
            close();
          }}
        />
      </div>

      <div
        className="tw-sheet"
        role="dialog"
        aria-label="Session skills"
        data-active={panel === 'skills' ? 'true' : 'false'}
        hidden={panel !== 'skills'}
      >
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

      <div
        className="tw-sheet"
        role="dialog"
        aria-label="Conversations"
        data-active={panel === 'sessions' ? 'true' : 'false'}
        hidden={panel !== 'sessions'}
      >
        <ConversationHistoryPanel
          conversations={conversations}
          disabled={busy}
          onRefresh={onRefreshConversations}
          onResume={onResumeConversation}
        />
      </div>
    </div>
  );
}
