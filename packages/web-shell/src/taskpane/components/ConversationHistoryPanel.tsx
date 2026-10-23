import { useState } from 'react';
import type { ConversationItem, ConversationsState } from '../../controller.js';
import { APP_NAME, appItem, appName } from './app-names.js';

export interface ConversationHistoryPanelProps {
  conversations: ConversationsState;
  /** The app this pane runs in; with "This app only" on, other apps' chats are hidden. */
  surface?: string;
  disabled?: boolean;
  onRefresh: () => void;
  onResume: (name: string) => void;
}

export function ConversationHistoryPanel({
  conversations,
  surface,
  disabled = false,
  onRefresh,
  onResume,
}: ConversationHistoryPanelProps): JSX.Element {
  const [thisAppOnly, setThisAppOnly] = useState(true);
  const filtering = Boolean(surface) && thisAppOnly;
  const items = filtering
    ? conversations.items.filter((item) => item.surface === surface)
    : conversations.items;
  const hiddenCount = conversations.items.length - items.length;
  // A chat from another app is continued only after the user confirms it here.
  const [confirming, setConfirming] = useState<ConversationItem>();
  const resume = (item: ConversationItem): void => {
    if (surface && item.surface && item.surface !== surface) setConfirming(item);
    else onResume(item.name);
  };
  return (
    <section className="sessions" aria-label="Conversations">
      <div className="sessions-head">
        <div>
          <span className="sessions-title">Conversations</span>
          <span className="sessions-meta">{sessionMeta(conversations, items.length)}</span>
        </div>
        <button
          type="button"
          className="mini-btn"
          disabled={disabled || conversations.loading}
          onClick={onRefresh}
        >
          {conversations.loading ? 'loading' : 'refresh'}
        </button>
      </div>

      {conversations.error && (
        <div className="sessions-error" role="status">
          {conversations.error}
        </div>
      )}

      {surface && (
        <label className="sessions-filter">
          <input
            type="checkbox"
            checked={thisAppOnly}
            onChange={(e) => setThisAppOnly(e.currentTarget.checked)}
          />
          {APP_NAME[surface] ?? surface} only
        </label>
      )}

      {!conversations.error && conversations.loaded && conversations.items.length === 0 && (
        <div className="sessions-empty">No conversations returned for this signed-in user.</div>
      )}

      {!conversations.error && conversations.loaded && items.length === 0 && hiddenCount > 0 && (
        <div className="sessions-empty">
          No {APP_NAME[surface ?? ''] ?? 'chats'} chats yet. Untick "{APP_NAME[surface ?? '']} only"
          to see {hiddenCount} from other apps.
        </div>
      )}

      {confirming && surface && confirming.surface ? (
        <div
          className="sessions-confirm"
          role="alertdialog"
          aria-labelledby="sessions-confirm-text"
        >
          <p id="sessions-confirm-text">
            This chat started in {appName(confirming.surface)}. Its answers are about that{' '}
            {appItem(confirming.surface)}. Continue it here in {appName(surface)}?
          </p>
          <div className="sessions-confirm-actions">
            <button
              type="button"
              className="session-resume"
              onClick={() => {
                setConfirming(undefined);
                onResume(confirming.name);
              }}
            >
              continue here
            </button>
            <button type="button" className="mini-btn" onClick={() => setConfirming(undefined)}>
              cancel
            </button>
          </div>
        </div>
      ) : null}

      <ol className="sessions-list">
        {items.map((item) => (
          <ConversationRow
            key={item.name}
            item={item}
            disabled={disabled || item.active}
            onResume={() => resume(item)}
          />
        ))}
      </ol>
    </section>
  );
}

function ConversationRow({
  item,
  disabled,
  onResume,
}: {
  item: ConversationItem;
  disabled: boolean;
  onResume: () => void;
}): JSX.Element {
  return (
    <li className={`session-row${item.active ? ' active' : ''}`}>
      <div className="session-main">
        <span className="session-title" title={item.title}>
          {item.title}
        </span>
        <span className="session-sub">
          {item.surface ? `${appName(item.surface)} · ` : ''}
          {item.turnCount} {item.turnCount === 1 ? 'turn' : 'turns'}
          {item.updatedAt ? ` · ${formatWhen(item.updatedAt)}` : ''}
          {item.isPinned ? ' · pinned' : ''}
        </span>
      </div>
      <button
        type="button"
        className="session-resume"
        disabled={disabled}
        aria-label={`Continue ${item.title}`}
        onClick={onResume}
      >
        {item.active ? 'active' : 'continue'}
      </button>
    </li>
  );
}

function sessionMeta(conversations: ConversationsState, shown: number): string {
  if (conversations.loading) return 'loading';
  if (!conversations.loaded) return 'not loaded';
  return `${shown} shown`;
}

function formatWhen(raw: string): string {
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return raw;
  return new Intl.DateTimeFormat(undefined, {
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  }).format(date);
}
