import type { QuickAction, Surface } from '@ge/contracts';
import { modeLabel } from './mode-labels.js';

/** "Lee, Ann" (directory order) and "Ann Lee" both greet "Ann"; an email-like name is skipped. */
export function firstName(displayName: string | undefined): string | undefined {
  const name = displayName?.trim();
  if (!name || name.includes('@')) return undefined;
  const given = name.includes(',') ? name.split(',')[1] : name;
  return given?.trim().split(/\s+/)[0] || undefined;
}

const ICON: Record<QuickAction['output'], JSX.Element> = {
  chat: (
    <>
      <circle cx="12" cy="12" r="8.5" />
      <path d="M9.6 9.6a2.4 2.4 0 1 1 3.4 2.2c-.6.3-1 .8-1 1.4v.6M12 16.6v.1" />
    </>
  ),
  annotation: (
    <>
      <path d="M4 6h16M4 10h16M4 14h10M4 18h7" />
    </>
  ),
  write: (
    <>
      <path d="M4 20h4L19 9l-4-4L4 16z" />
      <path d="M13 7l4 4" />
    </>
  ),
};

export function WorkspaceHome({
  userName,
  actions,
  disabled,
  onAction,
}: {
  surface: Surface;
  userName?: string;
  actions: QuickAction[];
  disabled: boolean;
  onAction: (action: QuickAction) => void;
}): JSX.Element {
  const name = firstName(userName);
  return (
    <section className="workspace-home" aria-label="Start a task">
      <h1>{name ? `Hi ${name}` : 'Hi'}</h1>
      <p>What can I help you with?</p>
      {actions.length > 0 && (
        <>
          <div className="workspace-try">Try these</div>
          <div className="workspace-actions">
            {actions.map((action) => (
              <button
                type="button"
                className="workspace-action"
                key={action.id}
                data-action-id={action.id}
                disabled={disabled}
                title={`${modeLabel(action)} · ${
                  action.ground.includes('unit') ? 'Attached sources' : 'Current context'
                }`}
                onClick={() => onAction(action)}
              >
                <svg className="workspace-action-icon" viewBox="0 0 24 24" aria-hidden="true">
                  {ICON[action.output]}
                </svg>
                <span>{action.label}</span>
              </button>
            ))}
          </div>
        </>
      )}
    </section>
  );
}
