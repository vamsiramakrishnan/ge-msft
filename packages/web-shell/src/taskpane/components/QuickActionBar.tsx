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
  const excluded = excludeIds ? new Set(excludeIds) : undefined;
  const actions = quickActionsForSurface(surface, allowedIntents).filter(
    (action) => !excluded?.has(action.id),
  );
  if (actions.length === 0) return null;

  return (
    <section className="quick-actions" aria-label="Quick actions">
      {actions.map((action) => (
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
          </span>
        </span>
      ))}
    </section>
  );
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
