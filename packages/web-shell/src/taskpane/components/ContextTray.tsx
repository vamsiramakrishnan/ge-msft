import type { ContextChip } from '../../controller.js';
import { useDisclosure } from './useDisclosure.js';

export interface ContextTrayProps {
  chips: ContextChip[];
  onToggle: (id: string, attach: boolean) => void;
  onRefresh: () => void;
}

const KIND_DOT: Readonly<Record<string, string>> = {
  document: 'var(--host)',
  selection: 'var(--host)',
  range: 'var(--host)',
  'mail-item': 'var(--link)',
  transcript: 'var(--teal)',
};

/**
 * The research-unit context tray: the grounding scope as removable chips. Attached chips carry an
 * × to detach; unattached chips are clickable to attach. "Refresh" re-reads what the host exposes
 * right now (selection/document/range/etc.). Matches the mockup's "Research unit · grounding scope".
 */
function Chip({
  chip,
  onToggle,
}: {
  chip: ContextChip;
  onToggle: (id: string, attach: boolean) => void;
}): JSX.Element {
  const dot = KIND_DOT[chip.kind] ?? 'var(--soft)';
  const detailId = `ctx-detail-${safeId(chip.id)}`;
  return (
    <span className="detail-hover chip-wrap">
      <span className={`chip${chip.attached ? ' on' : ''}`} aria-describedby={detailId}>
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
      <span id={detailId} className="detail-popover" role="tooltip">
        <strong>{chip.title}</strong>
        <span>{chip.kind}</span>
        {chip.preview ? <span>{chip.preview}</span> : null}
      </span>
    </span>
  );
}

export function ContextTray({ chips, onToggle, onRefresh }: ContextTrayProps): JSX.Element {
  // Read-only view shaping: surface attached sources first so the active grounding scope reads
  // top-to-bottom. The controller still owns the flat `chips` list and the `attached` flag.
  const attached = chips.filter((c) => c.attached);
  const available = chips.filter((c) => !c.attached);
  const preview = attached.length > 0 ? attached : available;
  const { open, toggle, containerRef } = useDisclosure<HTMLElement>();
  return (
    <section
      className="unit"
      aria-label="Research unit grounding scope"
      ref={containerRef}
      data-open={open ? 'true' : 'false'}
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
        <button
          type="button"
          className="unit-toggle"
          aria-expanded={open}
          aria-controls="ctx-chips"
          aria-label={open ? 'Hide context sources' : 'Show context sources'}
          onClick={toggle}
        >
          <span className="tw-caret" aria-hidden="true">
            ▾
          </span>
        </button>
      </div>
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
      <div id="ctx-chips" className="chips context-popover">
        {chips.length === 0 && (
          <span className="muted small">Nothing to attach yet. Refresh scans the host.</span>
        )}
        {attached.length > 0 && (
          <div className="chips-group" role="list" aria-label="Attached sources">
            {attached.map((chip) => (
              <span role="listitem" key={chip.id} style={{ display: 'contents' }}>
                <Chip chip={chip} onToggle={onToggle} />
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
                <Chip chip={chip} onToggle={onToggle} />
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
