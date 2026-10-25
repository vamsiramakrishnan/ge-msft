import { useId, useState } from 'react';
import type { RunStep } from '../../controller.js';

export interface RunStepsProps {
  steps: RunStep[];
  /** Builds the content-free support snapshot copied by "Copy diagnostics". */
  diagnostics?: () => string;
}

/**
 * The command-loop transcript: a compact, ordered list of the `CommandLoopEvent`s
 * (turn / command / read / write / done) so the user can watch the read-many/write-one loop
 * progress. Rendered as an ordered list with a polite live region so new steps are announced.
 */
export function RunSteps({ steps, diagnostics }: RunStepsProps): JSX.Element | null {
  const [expanded, setExpanded] = useState(false);
  const [copyState, setCopyState] = useState('');
  const [manualCopy, setManualCopy] = useState<string>();
  const listId = useId();
  const copyDiagnostics = async (): Promise<void> => {
    if (!diagnostics) return;
    const text = diagnostics();
    setManualCopy(undefined);
    if (await copyText(text)) {
      setCopyState('Copied. Paste it to your support contact.');
      return;
    }
    setManualCopy(text);
    setCopyState('Select the text below and copy it (Ctrl+C / ⌘C).');
  };

  if (steps.length === 0) return null;

  const latest = steps[steps.length - 1];
  const stepCount = steps.length === 1 ? '1 step' : `${steps.length} steps`;

  return (
    <section
      className={`run-steps${expanded ? ' is-expanded' : ''}`}
      aria-label="Command loop steps"
      aria-live="polite"
    >
      <button
        type="button"
        className="run-steps-toggle"
        aria-expanded={expanded}
        aria-controls={listId}
        onClick={() => setExpanded((value) => !value)}
      >
        <span className="run-steps-title eyebrow">Activity</span>
        <span className="run-steps-summary">
          <span className="run-steps-count">{stepCount}</span>
          {latest ? (
            <span className="run-steps-latest">
              <span className="step-kind">{latest.kind.replace(/-/g, ' ')}</span>
              <span className="step-text">{latest.text}</span>
            </span>
          ) : null}
        </span>
        <span className="run-steps-caret" aria-hidden="true" />
      </button>
      <ol
        id={listId}
        className="run-steps-list"
        hidden={!expanded}
        style={{ listStyle: 'none', margin: 0, padding: 0 }}
      >
        {steps.map((s) => (
          <li key={s.id} className={`run-step step-${s.kind}`}>
            <span className="step-kind">{s.kind.replace(/-/g, ' ')}</span>
            <span className="step-body">
              <span className="step-text">{s.text}</span>
              {s.artifact ? <WorkspaceArtifactCard artifact={s.artifact} /> : null}
            </span>
          </li>
        ))}
      </ol>
      {expanded && diagnostics ? (
        <div className="run-steps-diagnostics">
          <button type="button" onClick={() => void copyDiagnostics()}>
            Copy diagnostics
          </button>
          {copyState ? <span>{copyState}</span> : null}
          {manualCopy ? (
            <textarea
              readOnly
              aria-label="Diagnostics"
              value={manualCopy}
              onFocus={(e) => e.currentTarget.select()}
              ref={(el) => el?.select()}
            />
          ) : null}
        </div>
      ) : null}
    </section>
  );
}

function WorkspaceArtifactCard({
  artifact,
}: {
  artifact: NonNullable<RunStep['artifact']>;
}): JSX.Element {
  return (
    <article className="workspace-artifact-card">
      <header className="workspace-artifact-head">
        <span className="workspace-artifact-title">{artifact.title}</span>
      </header>
      {artifact.meta.length > 0 ? (
        <ul className="workspace-artifact-meta" aria-label="Artifact metadata">
          {artifact.meta.map((item) => (
            <li key={item}>{item}</li>
          ))}
        </ul>
      ) : null}
      {artifact.preview ? (
        <pre className="workspace-artifact-preview">{artifact.preview}</pre>
      ) : null}
      {artifact.matches && artifact.matches.length > 0 ? (
        <ol className="workspace-artifact-matches" aria-label="Artifact matches">
          {artifact.matches.map((match) => (
            <li key={`${match.line}:${match.text}`}>
              <span className="workspace-artifact-line">L{match.line}</span>
              <span>{match.text}</span>
            </li>
          ))}
        </ol>
      ) : null}
    </article>
  );
}

/** `document.execCommand('copy')` through a hidden textarea; false when the host refuses it. */
/**
 * Copy text, trying the legacy copy first: it runs synchronously inside the click, which Office on
 * Mac (Safari web view) requires, and Office desktop blocks the async clipboard API. The async API
 * is the fallback for browsers that disable the legacy command.
 */
export async function copyText(text: string, html?: string): Promise<boolean> {
  if (legacyCopy(text, html)) return true;
  try {
    if (html && typeof ClipboardItem !== 'undefined' && navigator.clipboard.write) {
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/plain': new Blob([text], { type: 'text/plain' }),
          'text/html': new Blob([html], { type: 'text/html' }),
        }),
      ]);
    } else {
      await navigator.clipboard.writeText(text);
    }
    return true;
  } catch {
    return false;
  }
}

/** The legacy copy command; with `html`, the clipboard also gets a formatted version. */
export function legacyCopy(text: string, html?: string): boolean {
  const focused = document.activeElement as HTMLElement | null;
  const onCopy = (event: ClipboardEvent): void => {
    if (!html || !event.clipboardData) return;
    event.clipboardData.setData('text/plain', text);
    event.clipboardData.setData('text/html', html);
    event.preventDefault();
  };
  document.addEventListener('copy', onCopy);
  const area = document.createElement('textarea');
  area.value = text;
  area.setAttribute('readonly', '');
  area.style.position = 'fixed';
  area.style.opacity = '0';
  document.body.appendChild(area);
  area.select();
  try {
    return document.execCommand('copy');
  } catch {
    return false;
  } finally {
    document.removeEventListener('copy', onCopy);
    area.remove();
    focused?.focus?.(); // back to the button that was clicked
  }
}
