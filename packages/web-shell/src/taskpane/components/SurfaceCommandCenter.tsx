import type { Intent, QuickAction, Surface } from '@ge/contracts';
import { quickActionsForSurface } from '@ge/contracts';

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
}: SurfaceCommandCenterProps): JSX.Element | null {
  const actions = surfacePrimaryActions(surface, allowedIntents);
  if (actions.length === 0) return null;
  const copy = SURFACE_COPY[surface];
  const status = hasGate ? 'Decision needed' : busy ? 'Working' : 'Ready';
  const state = hasGate ? 'gate' : busy ? 'busy' : 'ready';

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

      <div className="detail-hover surface-details">
        <button type="button" className="surface-detail-trigger" aria-describedby="surface-details">
          <span>{attachedCount} attached</span>
          <span>{availableCount} nearby</span>
          <span>{messageCount} turns</span>
        </button>
        <div id="surface-details" className="detail-popover surface-details-popover" role="tooltip">
          <div className="surface-metrics" aria-label="Current work state">
            <span>
              <strong>{attachedCount}</strong> attached
            </span>
            <span>
              <strong>{availableCount}</strong> nearby
            </span>
            <span>
              <strong>{messageCount}</strong> turns
            </span>
            <span>
              <strong>{proposalCount}</strong> proposals
            </span>
          </div>

          <div className="surface-entrypoints" aria-label="Ways to use Gemini in this host">
            {entrypoints(copy.object).map((entry) => (
              <span key={entry.id} className="detail-hover surface-entry-wrap">
                <span className="surface-entry" tabIndex={0} aria-describedby={`entry-${entry.id}`}>
                  <span className="surface-entry-dot" aria-hidden="true" />
                  <span>{entry.label}</span>
                </span>
                <span
                  id={`entry-${entry.id}`}
                  className="detail-popover surface-entry-detail"
                  role="tooltip"
                >
                  <strong>{entry.title}</strong>
                  <span>{entry.detail}</span>
                </span>
              </span>
            ))}
          </div>
        </div>
      </div>

      <div className="surface-primary-actions">
        {actions.map((action) => (
          <span key={action.id} className="detail-hover surface-action-wrap">
            <button
              type="button"
              className="surface-action"
              data-action-id={action.id}
              data-output={action.output}
              data-intent={action.intent}
              disabled={busy || hasGate}
              aria-describedby={`surface-action-detail-${action.id}`}
              onClick={() => onAction(action)}
            >
              <span className="surface-action-mode" data-output={action.output} aria-hidden="true">
                {shortMode(action)}
              </span>
              <span className="surface-action-body">
                <span className="surface-action-label">{action.label}</span>
                <span className="surface-action-meta">
                  {action.intent} · {action.scope.kind} · {actionMode(action)}
                </span>
              </span>
            </button>
            <span
              id={`surface-action-detail-${action.id}`}
              className="detail-popover surface-action-detail"
              role="tooltip"
            >
              <strong>{action.label}</strong>
              <span>{action.prompt}</span>
              <span>{actionMode(action)}</span>
              {action.contextMenu ? <span>Also available from the host context menu.</span> : null}
            </span>
          </span>
        ))}
      </div>
    </section>
  );
}

function entrypoints(
  object: string,
): { id: string; label: string; title: string; detail: string }[] {
  return [
    {
      id: 'pane',
      label: 'Pane',
      title: 'Task pane',
      detail: `Use chat, approvals, diagnostics, and grounded actions beside the active ${object}.`,
    },
    {
      id: 'slash',
      label: '/ verbs',
      title: 'Command palette',
      detail: 'Type / to reveal only the commands this host and profile can actually run.',
    },
    {
      id: 'context',
      label: 'Right-click',
      title: 'Context menu',
      detail: `Ask about the current selection without copying content out of the ${object}.`,
    },
    {
      id: 'ribbon',
      label: 'Ribbon',
      title: 'Office command surface',
      detail: 'Open the pane or trigger installed function commands from the add-in ribbon group.',
    },
  ];
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

function shortMode(action: QuickAction): string {
  switch (action.output) {
    case 'chat':
      return 'Ask';
    case 'annotation':
      return 'Review';
    case 'write':
      return 'Change';
  }
}

function actionMode(action: QuickAction): string {
  switch (action.output) {
    case 'chat':
      return 'Answer';
    case 'annotation':
      return 'Review gate';
    case 'write':
      return 'Preview gate';
  }
}
