import { useId, useMemo, useState } from 'react';
import { quickActionsForSurface, type Intent, type QuickAction, type Surface } from '@ge/contracts';
import { modeCta, modeLabel } from './mode-labels.js';

const PIN_KEY = 'ge.action-pins.v1';

/** Only catalog ids persist. Prompts, source names, and user-entered values never do. */
function readPins(): string[] {
  try {
    const parsed: unknown = JSON.parse(localStorage.getItem(PIN_KEY) ?? '[]');
    return Array.isArray(parsed)
      ? parsed.filter((id): id is string => typeof id === 'string').slice(0, 100)
      : [];
  } catch {
    return [];
  }
}

export function filterActions(
  actions: QuickAction[],
  query: string,
  filter: string,
  pinned: string[],
): QuickAction[] {
  const terms = query.toLocaleLowerCase().trim().split(/\s+/).filter(Boolean);
  return actions
    .filter((action) => {
      if (filter === 'pinned' && !pinned.includes(action.id)) return false;
      if (['chat', 'write', 'annotation'].includes(filter) && action.output !== filter)
        return false;
      const searchable =
        `${action.label} ${action.prompt} ${action.intent} ${action.scope.kind}`.toLocaleLowerCase();
      return terms.every((term) => searchable.includes(term));
    })
    .sort((a, b) => Number(pinned.includes(b.id)) - Number(pinned.includes(a.id)));
}

export function ActionLibrary({
  surface,
  allowedIntents,
  disabled,
  onAction,
}: {
  surface: Surface;
  allowedIntents?: Iterable<Intent>;
  disabled: boolean;
  onAction: (action: QuickAction) => void;
}): JSX.Element {
  const [query, setQuery] = useState('');
  const [filter, setFilter] = useState('all');
  const [pinned, setPinned] = useState(readPins);
  const id = useId();
  const actions = useMemo(
    () => quickActionsForSurface(surface, allowedIntents),
    [surface, allowedIntents],
  );
  const results = filterActions(actions, query, filter, pinned);
  const togglePin = (actionId: string): void => {
    const next = pinned.includes(actionId)
      ? pinned.filter((p) => p !== actionId)
      : [...pinned, actionId].slice(-100);
    setPinned(next);
    try {
      localStorage.setItem(PIN_KEY, JSON.stringify(next));
    } catch {
      /* Session-only when storage is unavailable. */
    }
  };
  return (
    <section className="action-library" aria-label="Action library">
      <label className="action-search" htmlFor={id}>
        <svg
          width="16"
          height="16"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.8"
          aria-hidden="true"
        >
          <circle cx="10.5" cy="10.5" r="6.5" />
          <path d="m16 16 5 5" />
        </svg>
        <input
          id={id}
          type="search"
          aria-label="Search actions"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder="Find a task, outcome, or command"
          autoComplete="off"
        />
      </label>
      <div className="library-filters" role="group" aria-label="Filter actions">
        {(
          [
            ['all', 'All'],
            ['pinned', 'Pinned'],
            ['chat', 'Ask'],
            ['annotation', 'Review'],
            ['write', 'Change'],
          ] as const
        ).map(([value, label]) => (
          <button
            type="button"
            key={value}
            aria-pressed={filter === value}
            onClick={() => setFilter(value)}
          >
            {label}
          </button>
        ))}
      </div>
      <p className="library-count" role="status">
        {results.length} {results.length === 1 ? 'action' : 'actions'} available
      </p>
      <div className="library-results">
        {results.map((action) => (
          <article className="library-action" key={action.id}>
            <div className="library-action-heading">
              <h3>{action.label}</h3>
              <button
                type="button"
                className="pin-action"
                aria-label={`${pinned.includes(action.id) ? 'Unpin' : 'Pin'} ${action.label}`}
                aria-pressed={pinned.includes(action.id)}
                onClick={() => togglePin(action.id)}
              >
                {pinned.includes(action.id) ? '★' : '☆'}
              </button>
            </div>
            <p>{action.prompt.replace(/\{\{([^}]+)\}\}/g, '…')}</p>
            <div className="library-action-foot">
              <span>
                {modeLabel(action)} · {action.scope.kind.replace(/-/g, ' ')}
              </span>
              <button
                type="button"
                className="library-run"
                disabled={disabled}
                data-action-id={action.id}
                onClick={() => onAction(action)}
              >
                {action.parameters?.length ? 'Configure' : modeCta(action)}{' '}
                <span aria-hidden="true">↗</span>
              </button>
            </div>
          </article>
        ))}
        {results.length === 0 && (
          <div className="library-empty">
            <strong>
              {filter === 'pinned' ? 'Keep useful actions close.' : 'No matching actions.'}
            </strong>
            <p>
              {filter === 'pinned'
                ? 'Pin actions from All to find them here.'
                : 'Try a shorter search or a different filter.'}
            </p>
            <button
              type="button"
              className="text-control"
              onClick={() => {
                setQuery('');
                setFilter('all');
              }}
            >
              Show all actions
            </button>
          </div>
        )}
      </div>
    </section>
  );
}
