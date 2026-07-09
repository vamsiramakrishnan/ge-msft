import type { ConversationItem, ConversationsState } from '../../controller.js';

export interface ConversationHistoryPanelProps {
  conversations: ConversationsState;
  disabled?: boolean;
  onRefresh: () => void;
  onResume: (name: string) => void;
}

export function ConversationHistoryPanel({
  conversations,
  disabled = false,
  onRefresh,
  onResume,
}: ConversationHistoryPanelProps): JSX.Element {
  return (
    <section className="sessions" aria-label="Conversations">
      <div className="sessions-head">
        <div>
          <span className="sessions-title">Conversations</span>
          <span className="sessions-meta">{sessionMeta(conversations)}</span>
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

      {!conversations.error && conversations.loaded && conversations.items.length === 0 && (
        <div className="sessions-empty">No conversations returned for this signed-in user.</div>
      )}

      <ol className="sessions-list">
        {conversations.items.map((item) => (
          <ConversationRow
            key={item.name}
            item={item}
            disabled={disabled || item.active}
            onResume={onResume}
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
  onResume: (name: string) => void;
}): JSX.Element {
  return (
    <li className={`session-row${item.active ? ' active' : ''}`}>
      <div className="session-main">
        <span className="session-title" title={item.title}>
          {item.title}
        </span>
        <span className="session-sub">
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
        onClick={() => onResume(item.name)}
      >
        {item.active ? 'active' : 'continue'}
      </button>
    </li>
  );
}

function sessionMeta(conversations: ConversationsState): string {
  if (conversations.loading) return 'loading';
  if (!conversations.loaded) return 'not loaded';
  return `${conversations.items.length} shown`;
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
