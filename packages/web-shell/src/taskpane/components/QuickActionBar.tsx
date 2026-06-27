import { useMemo, useState } from 'react';
import type { Surface, Intent } from '@ge/contracts';
import { quickActionsForSurface, type QuickAction } from '@ge/contracts';

export interface QuickActionBarProps {
  /** The current surface — scopes which actions are offered (capability closure, ADR-0006). */
  surface: Surface;
  /** When given, narrows the offered actions to those whose `intent` the surface can run. */
  allowedIntents?: Iterable<Intent>;
  /** Disable the chips while a turn is in flight. */
  busy: boolean;
  /** Hide actions already promoted into the contextual command center. */
  excludeIds?: Iterable<string>;
  /** Dispatch a chosen action. The parent routes `chat` → `send`, `write`/`annotation` → the
   *  plan/approval loop (`runCommands`) — this component never chooses the gate itself. */
  onAction: (action: QuickAction) => void;
}

/**
 * The quick-action chip row — the curated, one-tap verb catalog for the current surface (ADR-0004
 * grammar, scoped by capability closure, ADR-0006). Renders {@link quickActionsForSurface} as
 * clickable chips above the composer. Clicking one hands the whole {@link QuickAction} back to the
 * parent, which seeds the matching path: read-only `chat` actions start a grounded turn, while
 * `write`/`annotation` actions go through the existing fail-closed plan/approval loop. The chip is
 * the human-friendly front of the same grammar the model emits — it adds no new gate of its own.
 */
export function QuickActionBar({
  surface,
  allowedIntents,
  busy,
  excludeIds,
  onAction,
}: QuickActionBarProps): JSX.Element | null {
  const actions = useMemo(() => {
    const excluded = excludeIds ? new Set(excludeIds) : undefined;
    return quickActionsForSurface(surface, allowedIntents).filter(
      (action) => !excluded?.has(action.id),
    );
  }, [allowedIntents, excludeIds, surface]);
  const [preferredGroup, setPreferredGroup] = useState<QuickAction['output']>('write');
  const [expanded, setExpanded] = useState(false);
  if (actions.length === 0) return null;
  const groups = groupActions(actions);
  const availableGroups = GROUPS.filter((group) => groups[group.output].length > 0);
  const activeGroup = availableGroups.some((group) => group.output === preferredGroup)
    ? preferredGroup
    : (availableGroups[0]?.output ?? 'chat');
  const activeActions = groups[activeGroup];
  const visibleActions = expanded ? activeActions : activeActions.slice(0, 4);
  const hiddenCount = Math.max(0, activeActions.length - visibleActions.length);

  return (
    <section className="quick-actions action-drawer" aria-label="More actions">
      <div className="action-drawer-head">
        <div className="action-drawer-copy">
          <div className="eyebrow">More actions</div>
          <div className="action-drawer-summary">
            {actions.length} available on this surface · scoped by host capability
          </div>
        </div>
        <div className="action-drawer-controls">
          <div className="action-tabs" role="tablist" aria-label="Action type">
            {availableGroups.map((group) => (
              <button
                key={group.output}
                type="button"
                role="tab"
                className="action-tab"
                aria-selected={activeGroup === group.output}
                data-selected={activeGroup === group.output ? 'true' : 'false'}
                onClick={() => {
                  setPreferredGroup(group.output);
                  setExpanded(false);
                }}
              >
                <span>{group.label}</span>
                <strong>{groups[group.output].length}</strong>
              </button>
            ))}
          </div>
          {activeActions.length > 4 ? (
            <button
              type="button"
              className="mini-btn action-more"
              onClick={() => setExpanded((value) => !value)}
              aria-expanded={expanded}
            >
              {expanded ? 'Less' : `+${hiddenCount}`}
            </button>
          ) : null}
        </div>
      </div>

      <div className="action-drawer-list" data-group={activeGroup}>
        {visibleActions.map((action) => (
          <span key={action.id} className="detail-hover quick-action-wrap">
            <button
              type="button"
              className="quick-action"
              data-action-id={action.id}
              data-output={action.output}
              data-intent={action.intent}
              disabled={busy}
              aria-describedby={`qa-detail-${action.id}`}
              onClick={() => onAction(action)}
            >
              <span className="quick-action-icon" aria-hidden="true">
                {actionIcon(action.output)}
              </span>
              <span className="quick-action-main">{action.label}</span>
              <span className="quick-action-meta">{actionMeta(action.output)}</span>
            </button>
            <span id={`qa-detail-${action.id}`} className="detail-popover" role="tooltip">
              <strong>{action.label}</strong>
              <span>{action.prompt}</span>
              <span>
                {action.intent} · {action.scope.kind} · {actionMeta(action.output)}
              </span>
              {action.contextMenu ? <span>Also available from the host context menu.</span> : null}
            </span>
          </span>
        ))}
      </div>
    </section>
  );
}

const GROUPS: { output: QuickAction['output']; label: string }[] = [
  { output: 'write', label: 'Change' },
  { output: 'annotation', label: 'Review' },
  { output: 'chat', label: 'Ask' },
];

function groupActions(actions: QuickAction[]): Record<QuickAction['output'], QuickAction[]> {
  return {
    chat: actions.filter((action) => action.output === 'chat'),
    annotation: actions.filter((action) => action.output === 'annotation'),
    write: actions.filter((action) => action.output === 'write'),
  };
}

function actionIcon(output: QuickAction['output']): string {
  switch (output) {
    case 'chat':
      return '?';
    case 'annotation':
      return '+';
    case 'write':
      return '>';
  }
}

function actionMeta(output: QuickAction['output']): string {
  switch (output) {
    case 'chat':
      return 'Ask';
    case 'annotation':
      return 'Review';
    case 'write':
      return 'Preview';
  }
}
