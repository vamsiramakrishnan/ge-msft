import type { QuickAction, Surface } from '@ge/contracts';
import { modeLabel } from './mode-labels.js';

const COPY: Record<Surface, { title: string; detail: string }> = {
  word: {
    title: 'Make the next draft better.',
    detail: 'Tighten the argument, check the evidence, or review the whole document.',
  },
  excel: {
    title: 'Find the story in the numbers.',
    detail: 'Investigate a range, build a chart, or turn a question into a formula.',
  },
  powerpoint: {
    title: 'Give the deck a clear point.',
    detail: 'Build from your sources, test the narrative, or prepare the talk track.',
  },
  onenote: {
    title: 'Turn research into something useful.',
    detail: 'Connect the evidence, compare sources, and write a cited synthesis.',
  },
  outlook: {
    title: 'Get to the next action.',
    detail: 'Catch up on the thread, prepare a reply, or extract commitments.',
  },
  teams: {
    title: 'Leave with decisions and owners.',
    detail: 'Catch up on the discussion and turn the transcript into an actionable recap.',
  },
};

export function WorkspaceHome({
  surface,
  actions,
  disabled,
  onAction,
}: {
  surface: Surface;
  actions: QuickAction[];
  disabled: boolean;
  onAction: (action: QuickAction) => void;
}): JSX.Element {
  return (
    <section className="workspace-home" aria-label="Start a task">
      <span className="workspace-kicker">Your work, in context</span>
      <h1>{COPY[surface].title}</h1>
      <p>{COPY[surface].detail}</p>
      <div className="workspace-actions">
        {actions.map((action, i) => (
          <button
            type="button"
            className="workspace-action"
            key={action.id}
            data-action-id={action.id}
            disabled={disabled}
            onClick={() => onAction(action)}
          >
            <span className="workspace-action-index" aria-hidden="true">
              0{i + 1}
            </span>
            <span>
              <strong>{action.label}</strong>
              <small>
                {modeLabel(action)} ·{' '}
                {action.ground.includes('unit') ? 'Attached sources' : 'Current context'}
              </small>
            </span>
            <span className="workspace-action-arrow" aria-hidden="true">
              ↗
            </span>
          </button>
        ))}
      </div>
      <div className="workspace-hint">Choose an action, or describe what you need below.</div>
    </section>
  );
}
