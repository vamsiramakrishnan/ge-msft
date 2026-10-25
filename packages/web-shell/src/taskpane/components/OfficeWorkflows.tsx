import { useId, useState } from 'react';
import type { Surface } from '@ge/contracts';
import type { PanelController, PanelState } from '../../controller.js';
import { copyText } from './RunSteps.js';

export function OfficeWorkflows({
  surface,
  state,
  controller,
  disabled,
}: {
  surface: Surface;
  state: PanelState;
  controller: PanelController;
  disabled: boolean;
}): JSX.Element {
  const id = useId();
  const [scope, setScope] = useState('Current selection and governing definitions');
  const [playbook, setPlaybook] = useState('');
  const [instructions, setInstructions] = useState('');
  const [targetId, setTargetId] = useState('');
  const [nextText, setNextText] = useState('');
  const [inspected, setInspected] = useState<string[]>([]);
  const [disposable, setDisposable] = useState(false);
  const [fixtureRange, setFixtureRange] = useState('');
  const [reportCopy, setReportCopy] = useState('');
  const commit =
    typeof __BUILD_ID__ === 'string'
      ? /^([a-f0-9]{40})(?:\s|$)/.exec(__BUILD_ID__)?.[1]
      : undefined;
  const lab =
    typeof window !== 'undefined' &&
    new URLSearchParams(window.location.search).get('office-lab') === '1';
  const targets = controller.managedTextTargets?.() ?? [];
  const canRefresh =
    state.officeCapabilities?.includes(
      surface === 'word' ? 'fill-content-control' : 'set-shape-text',
    ) === true;
  const target = targets.find((ref) => ref.id === targetId);
  const unresolved =
    state.taskJournal?.entries.filter((entry) =>
      ['uncertain', 'dispatching'].includes(entry.state),
    ) ?? [];
  return (
    <>
      {lab && ['word', 'excel', 'powerpoint', 'outlook'].includes(surface) && (
        <details className="office-workbench">
          <summary>Native Office certification</summary>
          <div className="office-workflow-body">
            <p>
              Run a small native fixture through the normal approval flow. This does not call the
              model. Use a disposable document or draft.
            </p>
            {!commit && <p>A clean build from an exact Git commit is required.</p>}
            <label>
              <input
                type="checkbox"
                checked={disposable}
                onChange={(event) => setDisposable(event.target.checked)}
                disabled={disabled}
              />
              This is a disposable file or draft.
            </label>
            {surface === 'excel' && (
              <label className="field">
                <span className="field-label">Disposable 2 × 2 destination</span>
                <input
                  className="field-control"
                  value={fixtureRange}
                  onChange={(event) => setFixtureRange(event.target.value)}
                  placeholder="Sheet1!A1:B2"
                  disabled={disabled}
                />
              </label>
            )}
            <button
              type="button"
              className="btn"
              disabled={
                disabled || !commit || !disposable || (surface === 'excel' && !fixtureRange.trim())
              }
              onClick={() => void controller.runNativeCertification(surface, commit!, fixtureRange)}
            >
              Review native fixture
            </button>
            {state.nativeCertification && (
              <>
                <p>
                  Result: {state.nativeCertification.status} ·{' '}
                  {state.nativeCertification.nativeSyncs} native syncs
                </p>
                <button
                  type="button"
                  className="btn"
                  onClick={() =>
                    void copyText(JSON.stringify([state.nativeCertification], null, 2)).then((ok) =>
                      setReportCopy(ok ? 'Report copied.' : 'Select and copy the report below.'),
                    )
                  }
                >
                  Copy native report
                </button>
                <p role="status">{reportCopy}</p>
                <textarea
                  className="field-control"
                  aria-label="Native certification report"
                  readOnly
                  rows={8}
                  value={JSON.stringify([state.nativeCertification], null, 2)}
                />
              </>
            )}
          </div>
        </details>
      )}
      {surface === 'word' && (
        <details className="office-workbench">
          <summary>Contract review</summary>
          <form
            className="office-workflow-body"
            onSubmit={(event) => {
              event.preventDefault();
              void controller.runOfficeWorkflow(
                { id: 'contract-review', scope, playbook },
                surface,
              );
            }}
          >
            <label className="field" htmlFor={`${id}-scope`}>
              <span className="field-label">Review scope</span>
              <input
                id={`${id}-scope`}
                className="field-control"
                value={scope}
                onChange={(event) => setScope(event.target.value)}
                maxLength={16000}
                disabled={disabled}
                required
              />
            </label>
            <label className="field" htmlFor={`${id}-playbook`}>
              <span className="field-label">Review playbook</span>
              <textarea
                id={`${id}-playbook`}
                className="field-control"
                value={playbook}
                onChange={(event) => setPlaybook(event.target.value)}
                maxLength={16000}
                rows={5}
                disabled={disabled}
                required
              />
            </label>
            <p>
              Review proposed tracked edits and an exception register before changing the document.
            </p>
            <button
              className="btn"
              type="submit"
              disabled={disabled || !scope.trim() || !playbook.trim()}
            >
              Prepare contract review
            </button>
          </form>
        </details>
      )}
      {surface === 'outlook' && (
        <details className="office-workbench">
          <summary>Response package</summary>
          <form
            className="office-workflow-body"
            onSubmit={(event) => {
              event.preventDefault();
              void controller.runOfficeWorkflow({ id: 'outlook-response', instructions }, surface);
            }}
          >
            <label className="field" htmlFor={`${id}-reply`}>
              <span className="field-label">Response instructions</span>
              <textarea
                id={`${id}-reply`}
                className="field-control"
                value={instructions}
                onChange={(event) => setInstructions(event.target.value)}
                maxLength={16000}
                rows={4}
                disabled={disabled}
                required
              />
            </label>
            <p>
              Ground the response in the current email, check commitments and attachments, and
              review the draft.
            </p>
            <button className="btn" type="submit" disabled={disabled || !instructions.trim()}>
              Prepare response
            </button>
          </form>
        </details>
      )}
      {(surface === 'word' || surface === 'powerpoint') && canRefresh && (
        <details className="office-workbench">
          <summary>Refresh managed text</summary>
          <form
            className="office-workflow-body"
            onSubmit={(event) => {
              event.preventDefault();
              if (target) void controller.refreshManagedText(target, nextText);
            }}
          >
            <label className="field" htmlFor={`${id}-target`}>
              <span className="field-label">Output to update</span>
              <select
                id={`${id}-target`}
                className="field-control"
                value={targetId}
                onChange={(event) => setTargetId(event.target.value)}
                disabled={disabled}
                required
              >
                <option value="">Choose an output</option>
                {targets.map((ref) => (
                  <option key={ref.id} value={ref.id}>
                    {ref.title}
                  </option>
                ))}
              </select>
            </label>
            <button
              className="btn"
              type="button"
              disabled={disabled}
              onClick={() => void controller.refreshContext()}
            >
              Refresh available outputs
            </button>
            <label className="field" htmlFor={`${id}-text`}>
              <span className="field-label">New generated text</span>
              <textarea
                id={`${id}-text`}
                className="field-control"
                value={nextText}
                onChange={(event) => setNextText(event.target.value)}
                maxLength={16000}
                rows={5}
                disabled={disabled}
                required
              />
            </label>
            <p>
              Choose a Word content control or a PowerPoint text shape. Review the first update.
              Later refreshes detect edits to previously verified generated text.
            </p>
            <button
              className="btn"
              type="submit"
              disabled={disabled || !target || !nextText.trim()}
            >
              Review refresh
            </button>
          </form>
        </details>
      )}
      {unresolved.length > 0 && (
        <section className="office-workbench" aria-label="Interrupted Office changes">
          <h3>Inspect interrupted changes</h3>
          <p>
            An earlier change has an unresolved outcome. Check the document before starting another
            write.
          </p>
          {unresolved.map((entry) => (
            <div className="office-workflow-body" key={entry.id}>
              <p>
                {entry.kind} · {entry.state} · {new Date(entry.updatedAt).toLocaleString()}
              </p>
              <label>
                <input
                  type="checkbox"
                  checked={inspected.includes(entry.id)}
                  disabled={disabled}
                  onChange={(event) =>
                    setInspected((ids) =>
                      event.target.checked
                        ? [...ids, entry.id]
                        : ids.filter((value) => value !== entry.id),
                    )
                  }
                />
                I stopped other running panes and inspected this change in Office.
              </label>
              <button
                className="btn"
                type="button"
                disabled={disabled || !inspected.includes(entry.id)}
                onClick={() => void controller.acknowledgeTaskJournal(entry.id)}
              >
                Record inspection
              </button>
            </div>
          ))}
        </section>
      )}
    </>
  );
}
