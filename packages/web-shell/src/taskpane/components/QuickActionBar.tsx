import type { Surface, Intent } from '@ge/contracts';
import { quickActionsForSurface, type QuickAction } from '@ge/contracts';

export interface QuickActionBarProps {
  /** The current surface — scopes which actions are offered (capability closure, ADR-0006). */
  surface: Surface;
  /** When given, narrows the offered actions to those whose `intent` the surface can run. */
  allowedIntents?: Iterable<Intent>;
  /** Disable the chips while a turn is in flight. */
  busy: boolean;
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
  onAction,
}: QuickActionBarProps): JSX.Element | null {
  const actions = quickActionsForSurface(surface, allowedIntents);
  if (actions.length === 0) return null;

  return (
    <section className="quick-actions" aria-label="Quick actions">
      {actions.map((action) => (
        <button
          key={action.id}
          type="button"
          className="quick-action"
          data-output={action.output}
          data-intent={action.intent}
          disabled={busy}
          title={action.prompt}
          onClick={() => onAction(action)}
        >
          {action.label}
        </button>
      ))}
    </section>
  );
}
