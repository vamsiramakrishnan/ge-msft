import type { Intent, QuickAction, Surface } from '@ge/contracts';
import { quickActionsForSurface } from '@ge/contracts';
import { MODE_LABEL } from './mode-labels.js';

export interface SurfaceCommandCenterProps {
  surface: Surface;
  allowedIntents?: Iterable<Intent>;
  busy: boolean;
  hasGate: boolean;
  attachedCount: number;
  availableCount: number;
  messageCount: number;
  proposalCount: number;
  onAction: (action: QuickAction) => void;
}

interface SurfaceCopy {
  title: string;
  subtitle: string;
  glyph: string;
  object: string;
}

const SURFACE_COPY: Record<Surface, SurfaceCopy> = {
  word: {
    title: 'Word workspace',
    subtitle: 'Redlines, comments, and document review',
    glyph: 'W',
    object: 'document',
  },
  excel: {
    title: 'Excel workspace',
    subtitle: 'Ranges, formulas, comments, and charts',
    glyph: 'X',
    object: 'workbook',
  },
  powerpoint: {
    title: 'PowerPoint workspace',
    subtitle: 'Slide drafting, narrative review, and deck summaries',
    glyph: 'P',
    object: 'deck',
  },
  onenote: {
    title: 'OneNote workspace',
    subtitle: 'Page synthesis and source collection',
    glyph: 'N',
    object: 'page',
  },
  outlook: {
    title: 'Outlook workspace',
    subtitle: 'Threads, action items, and reviewable reply drafts',
    glyph: 'O',
    object: 'message',
  },
  teams: {
    title: 'Teams workspace',
    subtitle: 'Transcript recap and meeting actions',
    glyph: 'T',
    object: 'meeting',
  },
};

const ACTION_PRIORITY: Record<Surface, string[]> = {
  word: ['tighten', 'comment-on-issues', 'review-against', 'exec-summary'],
  excel: [
    'create-chart',
    'summarize-range',
    'find-anomalies',
    'risk-column',
    'write-formula',
    'explain-formula',
  ],
  powerpoint: ['draft-slide', 'draft-section', 'redesign', 'speaker-notes', 'summarize-deck'],
  onenote: ['synthesize-page', 'discover-sources', 'audio-overview', 'add-sources-to-unit'],
  outlook: ['draft-reply', 'catch-up', 'extract-actions', 'summarize-email', 'draft-new-email'],
  teams: ['live-notes', 'action-items', 'catch-up-teams'],
};

export function surfacePrimaryActions(
  surface: Surface,
  allowedIntents?: Iterable<Intent>,
): QuickAction[] {
  const priority = ACTION_PRIORITY[surface];
  const rank = new Map(priority.map((id, index) => [id, index]));
  return [...quickActionsForSurface(surface, allowedIntents)]
    .sort((a, b) => {
      const ar = rank.get(a.id) ?? 100;
      const br = rank.get(b.id) ?? 100;
      if (ar !== br) return ar - br;
      return outputRank(a) - outputRank(b);
    })
    .slice(0, 3);
}

/**
 * The surface command strip: who the host is, whether the pane is ready, the three highest-value
 * actions for this host as visible chips, and a one-line grounding/turn summary. The full action
 * catalog lives in the quick-action drawer; catalog routing lives behind the header settings. Keeping
 * the primary actions VISIBLE (not behind a disclosure) is the point — they are the task accelerators,
 * and hiding them wastes the strip. Everything else stays one compact line so the thread dominates.
 */
export function SurfaceCommandCenter({
  surface,
  allowedIntents,
  busy,
  hasGate,
  attachedCount,
  availableCount,
  messageCount,
  proposalCount,
  onAction,
}: SurfaceCommandCenterProps): JSX.Element {
  const actions = surfacePrimaryActions(surface, allowedIntents);
  const copy = SURFACE_COPY[surface];
  const status = hasGate ? 'Decision needed' : busy ? 'Working' : 'Ready';
  const state = hasGate ? 'gate' : busy ? 'busy' : 'ready';
  const disabled = busy || hasGate;

  return (
    <section className="surface-center" aria-label={`${copy.title} command center`}>
      <div className="surface-center-head">
        <div className="surface-identity">
          <span className="surface-glyph" aria-hidden="true">
            {copy.glyph}
          </span>
          <div className="surface-copy">
            <div className="surface-title">{copy.title}</div>
            <div className="surface-subtitle">{copy.subtitle}</div>
          </div>
        </div>
        <span className="surface-state" data-state={state}>
          {status}
        </span>
      </div>

      {actions.length > 0 ? (
        <div
          className="surface-primary-actions"
          role="group"
          aria-label={`Primary actions for ${copy.title}`}
        >
          {actions.map((action) => (
            <button
              key={action.id}
              type="button"
              className="surface-action"
              data-action-id={action.id}
              data-output={action.output}
              data-intent={action.intent}
              disabled={disabled}
              title={action.prompt}
              onClick={() => onAction(action)}
            >
              <span className="surface-action-mode" data-output={action.output} aria-hidden="true">
                {MODE_LABEL[action.output]}
              </span>
              <span className="surface-action-label">{action.label}</span>
            </button>
          ))}
        </div>
      ) : (
        <span className="surface-empty">No quick actions for this host yet.</span>
      )}

      <div className="surface-metrics" aria-label="Current work state">
        {attachedCount} attached · {availableCount} nearby · {messageCount} turns · {proposalCount}{' '}
        proposals
      </div>
    </section>
  );
}

function outputRank(action: QuickAction): number {
  switch (action.output) {
    case 'write':
      return 0;
    case 'annotation':
      return 1;
    case 'chat':
      return 2;
  }
}
