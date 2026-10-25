import { useId, useState } from 'react';
import type { PanelController, PanelState } from '../../controller.js';

/** Saved commands are inspected as data. The runtime owns every validation and execution decision. */
export function SavedOfficeTasks({
  state,
  controller,
  disabled,
}: {
  state: PanelState;
  controller: PanelController;
  disabled: boolean;
}): JSX.Element | null {
  const id = useId();
  const [name, setName] = useState('');
  const [commands, setCommands] = useState('');
  const [literal, setLiteral] = useState(false);
  const tasks = state.savedOfficeTasks;
  if (!tasks?.available) return null;
  const locked = disabled || tasks.loading;
  const selected = tasks.selected;
  const canArchive =
    selected?.record.stages.every((stage) => stage.state === 'verified') === true &&
    !selected.record.runId;
  const hasRemaining = selected?.record.stages.some((stage) => stage.state !== 'verified') === true;
  return (
    <details className="office-workbench saved-office-tasks">
      <summary>Saved tasks</summary>
      <div className="office-workflow-body">
        <p>
          Tasks belong to the current saved document and signed-in account. Nothing resumes
          automatically.
        </p>
        <button
          className="btn"
          type="button"
          disabled={locked}
          onClick={() => void controller.refreshSavedOfficeTasks()}
        >
          {tasks.loading ? 'Checking tasks…' : 'Refresh tasks'}
        </button>
        {tasks.error && (
          <p role="alert" className="saved-office-task-error">
            {tasks.error}
          </p>
        )}
        {!tasks.loaded && !tasks.loading && !tasks.error && (
          <p>Refresh to inspect tasks saved for this document.</p>
        )}
        {tasks.loaded && tasks.items.length === 0 && <p>No tasks are saved for this document.</p>}
        {tasks.items.length > 0 && (
          <ul className="saved-office-task-list">
            {tasks.items.map((task) => (
              <li key={task.id}>
                <div>
                  <strong>{task.id}</strong>
                  <span>
                    {task.verifiedStages} of {task.stageCount} stages verified · {task.state}
                  </span>
                  <span>
                    {task.durable
                      ? 'Stored on this device'
                      : 'Memory only; closing loses this record'}
                  </span>
                </div>
                <button
                  className="btn sm"
                  type="button"
                  disabled={locked}
                  onClick={() => void controller.inspectSavedOfficeTask(task.id)}
                >
                  Inspect task
                </button>
              </li>
            ))}
          </ul>
        )}
        {selected && (
          <section
            className="saved-office-task-inspection"
            aria-label={`Inspection of ${selected.record.program.id}`}
          >
            <h3>{selected.record.program.id}</h3>
            <p>Native sources and recorded outputs are checked again before remaining steps run.</p>
            <ol className="saved-office-task-stages">
              {selected.stages.map((stage) => (
                <li key={stage.id}>
                  <strong>
                    {stage.id}: {stage.status}
                  </strong>
                  {stage.reason && <span>{stage.reason}</span>}
                </li>
              ))}
            </ol>
            {selected.blockers.length > 0 && (
              <div className="saved-office-task-warning" role="status">
                <strong>Review is blocked</strong>
                <ul>
                  {selected.blockers.map((reason, index) => (
                    <li key={index}>{reason}</li>
                  ))}
                </ul>
              </div>
            )}
            {!hasRemaining && (
              <p>All recorded stages are verified. Inspection does not replay them.</p>
            )}
            <details>
              <summary>Inspect exact commands</summary>
              {selected.record.program.stages.map((stage) => (
                <div key={stage.id}>
                  <p>{stage.id}</p>
                  <pre>{stage.program}</pre>
                </div>
              ))}
            </details>
            <div className="saved-office-task-actions">
              <button
                className="btn pr"
                type="button"
                disabled={locked || !selected.canResume || !hasRemaining}
                onClick={() => void controller.resumeSavedOfficeTask(selected.record.program.id)}
              >
                Review remaining
              </button>
              <button
                className="btn danger"
                type="button"
                disabled={locked || !canArchive}
                onClick={() => void controller.archiveSavedOfficeTask(selected.record.program.id)}
              >
                Hide completed task
              </button>
            </div>
            <p>
              Hiding removes a completed task from this list. Its safety history stays recorded to
              prevent duplicate execution.
            </p>
          </section>
        )}
        <details className="saved-office-task-create">
          <summary>Save exact commands</summary>
          <form
            onSubmit={(event) => {
              event.preventDefault();
              if (!locked && literal)
                void controller.saveLiteralOfficeTask(name, commands, literal);
            }}
          >
            <p>Save one deterministic host change for later review. Saving does not execute it.</p>
            <p>
              Supported commands write literal values into bounded cells or managed text objects.
            </p>
            <label className="field" htmlFor={`${id}-name`}>
              <span className="field-label">Task ID</span>
              <input
                id={`${id}-name`}
                className="field-control"
                value={name}
                onChange={(event) => setName(event.target.value)}
                pattern="[a-z][a-z0-9-]{0,63}"
                maxLength={64}
                placeholder="month-close-memo"
                required
                disabled={locked}
              />
            </label>
            <label className="field" htmlFor={`${id}-commands`}>
              <span className="field-label">Exact commands</span>
              <textarea
                id={`${id}-commands`}
                className="field-control"
                rows={6}
                value={commands}
                onChange={(event) => {
                  setCommands(event.target.value);
                  setLiteral(false);
                }}
                maxLength={65536}
                required
                disabled={locked}
              />
            </label>
            <label className="saved-office-task-attestation">
              <input
                type="checkbox"
                checked={literal}
                onChange={(event) => setLiteral(event.target.checked)}
                disabled={locked}
              />
              These are my literal instructions. They do not depend on uncaptured document evidence.
            </label>
            <button
              className="btn"
              type="submit"
              disabled={locked || !literal || !name.trim() || !commands.trim()}
            >
              Save and inspect
            </button>
          </form>
        </details>
      </div>
    </details>
  );
}
