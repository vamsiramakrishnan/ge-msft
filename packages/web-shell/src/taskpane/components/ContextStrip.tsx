import { useState } from 'react';
import type { ContextChip } from '../../controller.js';

interface ContextStripProps {
  chips: ContextChip[];
  disabled: boolean;
  onToggle: (id: string, attached: boolean) => void;
  onReveal: (id: string) => void;
  onRefresh: () => void;
}

/** Always-visible source controls. The controller remains the sole attachment authority. */
export function ContextStrip({
  chips,
  disabled,
  onToggle,
  onReveal,
  onRefresh,
}: ContextStripProps): JSX.Element {
  const [expanded, setExpanded] = useState(false);
  const attached = chips.filter((chip) => chip.attached);
  const nearby = chips.filter((chip) => !chip.attached);
  const visible = expanded ? chips : attached.slice(0, 3);
  return (
    <section className="context-strip" aria-label="Active context">
      <div className="context-strip-heading">
        <span>
          Attached context ·{' '}
          <strong>
            {attached.length} {attached.length === 1 ? 'source' : 'sources'}
          </strong>
        </span>
        <button
          type="button"
          className="text-control"
          aria-expanded={expanded}
          aria-controls="active-context-chips"
          onClick={() => setExpanded(!expanded)}
        >
          {expanded ? 'Done' : nearby.length ? `+ ${nearby.length} available` : 'Manage'}
        </button>
      </div>
      <div id="active-context-chips" className="smart-chips">
        {visible.map((chip) => (
          <span className="smart-chip" data-attached={chip.attached} key={chip.id}>
            <button
              type="button"
              className="smart-chip-title"
              disabled={disabled || (chip.attached && !chip.revealable)}
              title={chip.preview ?? chip.title}
              aria-label={chip.attached ? `Locate ${chip.title}` : `Use ${chip.title}`}
              onClick={() => (chip.attached ? onReveal(chip.id) : onToggle(chip.id, true))}
            >
              <span className="smart-chip-symbol" aria-hidden="true">
                {chip.attached ? '◈' : '+'}
              </span>
              <span>{chip.title}</span>
            </button>
            {chip.attached && (
              <button
                type="button"
                className="smart-chip-remove"
                disabled={disabled}
                aria-label={`Remove ${chip.title} from context`}
                onClick={() => onToggle(chip.id, false)}
              >
                ×
              </button>
            )}
          </span>
        ))}
        {!expanded && attached.length > 3 && (
          <button type="button" className="text-control" onClick={() => setExpanded(true)}>
            +{attached.length - 3} more
          </button>
        )}
        {attached.length === 0 && !expanded && (
          <span className="context-empty">Add sources to focus the answer.</span>
        )}
        {expanded && (
          <button type="button" className="text-control" disabled={disabled} onClick={onRefresh}>
            Refresh sources
          </button>
        )}
      </div>
    </section>
  );
}
