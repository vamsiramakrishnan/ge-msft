import { useState } from 'react';
import type { SavedDiagnostic } from '../../saved-diagnostics.js';
import { copyText } from './RunSteps.js';

export interface SavedDiagnosticsPanelProps {
  /** Newest first. */
  entries: SavedDiagnostic[];
  onClear?: () => void;
}

/**
 * Diagnostics of earlier requests, saved on this device so they survive closing the pane. Each copy
 * is `{ question, diagnostic }`, ready to paste to support.
 */
export function SavedDiagnosticsPanel({
  entries,
  onClear,
}: SavedDiagnosticsPanelProps): JSX.Element {
  const [copied, setCopied] = useState<{ id: string; ok: boolean }>();
  const copy = async (entry: SavedDiagnostic): Promise<void> => {
    const text = JSON.stringify(
      { question: entry.question, diagnostic: entry.diagnostic },
      null,
      2,
    );
    setCopied({ id: entry.id, ok: await copyText(text) });
  };
  return (
    <section className="sessions" aria-label="Earlier diagnostics">
      <p className="saved-diagnostics-intro">
        Copy the diagnostics of a request and send them to support. Saved on this device for 7 days.
      </p>
      <ol className="sessions-list">
        {entries.map((entry) => (
          <li className="session-row" key={entry.id}>
            <div className="session-main">
              <span className="session-title" title={entry.question}>
                {entry.question}
              </span>
              <span className="session-sub">{formatWhen(entry.at)}</span>
            </div>
            <button
              type="button"
              className="session-resume"
              aria-label={`Copy diagnostics for ${entry.question}`}
              onClick={() => void copy(entry)}
            >
              {copied?.id === entry.id ? (copied.ok ? 'copied' : 'copy failed') : 'copy'}
            </button>
          </li>
        ))}
      </ol>
      {onClear && (
        <button type="button" className="mini-btn saved-diagnostics-clear" onClick={onClear}>
          clear saved diagnostics
        </button>
      )}
    </section>
  );
}

function formatWhen(raw: string): string {
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return raw;
  return new Intl.DateTimeFormat(undefined, {
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  }).format(date);
}
