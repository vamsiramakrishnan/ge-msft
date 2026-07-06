import type { ContextChip } from '../../controller.js';

export interface ContextTrayProps {
  chips: ContextChip[];
  onToggle: (id: string, attach: boolean) => void;
  onReveal: (id: string) => void;
  onRefresh: () => void;
  /**
   * Render expanded content inline (no self-collapse, no peek/toggle) for use inside the toolbar's
   * icon-triggered sheet — the sheet itself is the disclosure, so a second one would double-gate.
   */
  embedded?: boolean;
}

/**
 * The research-unit context tray: the grounding scope as removable chips. Attached chips carry an
 * × to detach; unattached chips are clickable to attach. "Refresh" re-reads what the host exposes
 * right now (selection/document/range/etc.). Matches the mockup's "Research unit · grounding scope".
 * The chip's dot is a status lamp styled entirely in CSS — blue attached, gray available — so
 * status is never color-alone-per-kind (Starlight spec: lamp + word).
 */
function Chip({
  chip,
  onToggle,
  onReveal,
}: {
  chip: ContextChip;
  onToggle: (id: string, attach: boolean) => void;
  onReveal: (id: string) => void;
}): JSX.Element {
  const detailId = `ctx-detail-${safeId(chip.id)}`;
  const revealLabel = `Open ${chip.title} in the host`;
  return (
    <span className="detail-hover chip-wrap">
      <span
        className={`chip${chip.attached ? ' on' : ''}`}
        aria-describedby={detailId}
        data-kind={chip.kind}
      >
        <span className="dot" aria-hidden="true" />
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
        {chip.revealable && (
          <button
            type="button"
            className="open-host"
            onClick={() => onReveal(chip.id)}
            aria-label={revealLabel}
            title={revealLabel}
          >
            ↗
          </button>
        )}
      </span>
      <span id={detailId} className="detail-popover" role="tooltip">
        <strong>{chip.title}</strong>
        <span>{chip.kind}</span>
        {chip.revealable ? <span>Open jumps to this location in the host.</span> : null}
        {chip.preview ? <span>{chip.preview}</span> : null}
      </span>
    </span>
  );
}

export function ContextTray({
  chips,
  onToggle,
  onReveal,
  onRefresh,
  embedded = false,
}: ContextTrayProps): JSX.Element {
  // Read-only view shaping: surface attached sources first so the active grounding scope reads
  // top-to-bottom. The controller still owns the flat `chips` list and the `attached` flag.
  const attached = chips.filter((c) => c.attached);
  const available = chips.filter((c) => !c.attached);
  const preview = attached.length > 0 ? attached : available;
  return (
    <section
      className={`unit detail-hover${embedded ? ' unit--embedded' : ''}`}
      aria-label="Research unit grounding scope"
    >
      <div className="unit-h">
        <span className="nb" aria-hidden="true" />
        <span className="unit-title">Context</span>
        <span className="unit-summary">
          {attached.length} attached · {available.length} nearby
        </span>
        <button
          type="button"
          className="link"
          onClick={onRefresh}
          aria-label="Refresh attachable context"
        >
          Refresh
        </button>
      </div>
      {!embedded && (
        <div className="unit-peek" aria-hidden="true">
          {chips.length === 0 ? (
            <span className="muted small">No host context loaded</span>
          ) : (
            preview.slice(0, 3).map((chip) => (
              <span key={chip.id} className={`mini-chip${chip.attached ? ' on' : ''}`}>
                {chip.title}
              </span>
            ))
          )}
        </div>
      )}
      <div className="chips context-popover">
        {chips.length === 0 && (
          <span className="muted small">Nothing to attach yet. Refresh scans the host.</span>
        )}
        {attached.length > 0 && (
          <div className="chips-group" role="list" aria-label="Attached sources">
            {attached.map((chip) => (
              <span role="listitem" key={chip.id} style={{ display: 'contents' }}>
                <Chip chip={chip} onToggle={onToggle} onReveal={onReveal} />
              </span>
            ))}
          </div>
        )}
        {attached.length > 0 && available.length > 0 && (
          <div className="chips-divider" aria-hidden="true" />
        )}
        {available.length > 0 && (
          <div className="chips-group" role="list" aria-label="Available to attach">
            {available.map((chip) => (
              <span role="listitem" key={chip.id} style={{ display: 'contents' }}>
                <Chip chip={chip} onToggle={onToggle} onReveal={onReveal} />
              </span>
            ))}
          </div>
        )}
      </div>
    </section>
  );
}

function safeId(value: string): string {
  return value.replace(/[^a-z0-9_-]/gi, '-');
}
