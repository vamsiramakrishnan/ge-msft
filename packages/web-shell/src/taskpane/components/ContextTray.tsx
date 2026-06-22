import type { ContextChip } from '../../controller.js';

export interface ContextTrayProps {
  chips: ContextChip[];
  onToggle: (id: string, attach: boolean) => void;
  onRefresh: () => void;
}

const KIND_DOT: Readonly<Record<string, string>> = {
  document: 'var(--grad)',
  selection: 'var(--grad)',
  range: 'var(--grad)',
  'mail-item': '#2f6fed',
  transcript: '#0f9d8c',
};

/**
 * The research-unit context tray: the grounding scope as removable chips. Attached chips carry an
 * × to detach; unattached chips are clickable to attach. "Refresh" re-reads what the host exposes
 * right now (selection/document/range/etc.). Matches the mockup's "Research unit · grounding scope".
 */
export function ContextTray({ chips, onToggle, onRefresh }: ContextTrayProps): JSX.Element {
  return (
    <section className="unit" aria-label="Research unit grounding scope">
      <div className="unit-h">
        <span className="nb" aria-hidden="true" />
        <span>Research unit · grounding scope</span>
        <button
          type="button"
          className="link"
          onClick={onRefresh}
          aria-label="Refresh attachable context"
        >
          Refresh
        </button>
      </div>
      <div className="chips">
        {chips.length === 0 && (
          <span className="muted">Nothing to attach yet — Refresh to scan the host.</span>
        )}
        {chips.map((chip) => {
          const dot = KIND_DOT[chip.kind] ?? '#9aa0b4';
          return (
            <span
              key={chip.id}
              className={`chip${chip.attached ? ' on' : ''}`}
              title={chip.preview ?? chip.title}
            >
              <span className="dot" style={{ background: dot }} aria-hidden="true" />
              {!chip.attached ? (
                <button
                  type="button"
                  className="chip-label"
                  onClick={() => onToggle(chip.id, true)}
                  aria-label={`Attach ${chip.title}`}
                >
                  {chip.title}
                </button>
              ) : (
                <span className="chip-label">{chip.title}</span>
              )}
              {chip.attached && (
                <button
                  type="button"
                  className="x"
                  onClick={() => onToggle(chip.id, false)}
                  aria-label={`Detach ${chip.title}`}
                >
                  ×
                </button>
              )}
            </span>
          );
        })}
      </div>
    </section>
  );
}
