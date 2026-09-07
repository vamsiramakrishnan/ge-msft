import { useEffect, useId, useMemo, useRef, useState } from 'react';
import {
  listWorkflowRecipes,
  type WorkflowRecipeDefinition,
  type WorkflowRecipeField,
} from '@ge/runtime';
import type { PanelController, PanelState } from '../../controller.js';
import { browserWorkflowPresets } from '../../workflow-preset-store.js';
import './workflow-workbench.css';

type Inputs = Record<string, unknown>;
type FormError = { field?: string; message: string };

function defaults(recipe: WorkflowRecipeDefinition): Inputs {
  return Object.fromEntries(
    recipe.fields
      .filter((field) => field.name !== 'destination' && field.default !== undefined)
      .map((field) => [field.name, field.default]),
  );
}

function inputKey(inputs: Inputs): string {
  return JSON.stringify(Object.entries(inputs).sort(([a], [b]) => a.localeCompare(b)));
}

function message(error: unknown): string {
  return error instanceof Error ? error.message : 'This action could not finish. Try again.';
}

function WorkflowIcon({ id }: { id: string }): JSX.Element {
  return (
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <g stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
        {id === 'reconcile-tables' ? (
          <>
            <path d="M4 7h13m-3-3 3 3-3 3M20 17H7m3-3-3 3 3 3" />
            <path d="M4 13v6m16-14v6" />
          </>
        ) : id === 'duplicate-rows' ? (
          <>
            <rect x="4" y="4" width="11" height="13" rx="2" />
            <path d="M18 8h2v12H9v-1M7 8h5m-5 4h5" />
          </>
        ) : (
          <>
            <path d="M4 20h16M6 16v-5m6 5V5m6 11V9" />
          </>
        )}
      </g>
    </svg>
  );
}

/** Workflow forms, previews and saved settings are derived from the runtime's versioned contracts. */
export function WorkflowWorkbench({
  state,
  controller,
  disabled,
}: {
  state: PanelState;
  controller: PanelController;
  disabled: boolean;
}): JSX.Element | null {
  const id = useId();
  const recipes = useMemo(() => listWorkflowRecipes(), []);
  const [recipeId, setRecipeId] = useState<string>();
  const recipe = recipes.find((item) => item.id === recipeId);
  const [inputs, setInputs] = useState<Inputs>({});
  const [destination, setDestination] = useState('');
  const [errors, setErrors] = useState<FormError[]>([]);
  const [localAction, setLocalAction] = useState<'preview' | 'write'>();
  const localBusy = localAction !== undefined;
  const [notice, setNotice] = useState('');
  const [staleResultId, setStaleResultId] = useState<string>();
  const [focusInputs, setFocusInputs] = useState(false);
  const sequence = useRef(0);
  const form = useRef<HTMLFormElement>(null);
  const errorPanel = useRef<HTMLDivElement>(null);
  const store = useMemo(browserWorkflowPresets, []);
  const [saved, setSaved] = useState(() => {
    try {
      return store?.list() ?? [];
    } catch {
      return [];
    }
  });
  const [storageError, setStorageError] = useState(() => {
    try {
      store?.list();
      return '';
    } catch {
      return 'Saved settings could not be loaded. You can clear them and save these settings again.';
    }
  });
  useEffect(() => {
    if (focusInputs) {
      form.current?.querySelector<HTMLInputElement>('input:not([type="checkbox"])')?.focus();
      setFocusInputs(false);
    }
  }, [focusInputs]);
  useEffect(() => {
    return () => {
      sequence.current += 1;
    };
  }, []);
  useEffect(() => {
    if (!errors.length) return;
    const field = errors.find((error) => error.field)?.field;
    const control = field ? form.current?.elements.namedItem(field) : undefined;
    if (control instanceof HTMLElement) {
      const advanced = control.closest('details');
      if (advanced) advanced.open = true;
      control.focus();
    } else errorPanel.current?.focus();
  }, [errors]);
  useEffect(() => {
    if (state.workflowRun?.recipeId === recipeId && state.workflowRun?.status === 'failed')
      errorPanel.current?.focus();
  }, [state.workflowRun?.runId, state.workflowRun?.status, state.workflowRun?.recipeId, recipeId]);

  if (!state.analysis || state.workflowRecipesAvailable === false) return null;
  const busy = disabled || localBusy;
  const parsed = recipe?.inputSchema.safeParse(inputs);
  const canonical = parsed?.success ? (parsed.data as Inputs) : undefined;
  const run = state.workflowRun?.recipeId === recipe?.id ? state.workflowRun : undefined;
  const matching = Boolean(
    run &&
    canonical &&
    recipe &&
    run.recipeVersion === recipe.version &&
    inputKey(run.inputs) === inputKey(canonical),
  );
  const artifact = run?.resultId
    ? state.analysis.artifacts.find((item) => item.id === run.resultId)
    : undefined;
  const stale = artifact?.id === staleResultId || Boolean(run?.error?.includes('Source changed:'));
  const result = matching && run?.status === 'ready' ? artifact : undefined;
  const changed = Boolean(run && !matching && run.status !== 'running');
  const write = run?.write;
  const sameDestination = write?.destination === destination.trim();
  const duplicateWrite =
    sameDestination && ['pending', 'written', 'uncertain'].includes(write?.status ?? '');
  const canWrite = Boolean(
    result &&
    result.rowCount > 0 &&
    !result.truncated &&
    !stale &&
    !busy &&
    state.workflowWritesAvailable !== false &&
    write?.status !== 'uncertain' &&
    destination.trim() &&
    !duplicateWrite,
  );
  const savedRecipe = saved.find((item) => item.recipeId === recipe?.id);
  const ranges = [
    ...new Set(
      state.analysis.artifacts.flatMap((item) => item.sources.map((source) => source.locator)),
    ),
  ].slice(0, 16);

  function choose(next: WorkflowRecipeDefinition): void {
    sequence.current += 1;
    setRecipeId(next.id);
    setInputs(defaults(next));
    setDestination('');
    setErrors([]);
    setNotice('');
    setFocusInputs(true);
  }

  function update(name: string, value: unknown): void {
    sequence.current += 1;
    setInputs((previous) => {
      const next = { ...previous };
      if (value === undefined) delete next[name];
      else next[name] = value;
      return next;
    });
    setErrors((previous) => previous.filter((error) => error.field !== name));
    setNotice('');
  }

  function validate(): Inputs | undefined {
    if (!recipe) return;
    const value = recipe.inputSchema.safeParse(inputs);
    if (!value.success) {
      setErrors(
        value.error.issues.map((issue) => ({
          field: String(issue.path[0] ?? ''),
          message: issue.message,
        })),
      );
      return;
    }
    setErrors([]);
    return value.data as Inputs;
  }

  async function preview(): Promise<void> {
    const values = validate();
    if (!values || !recipe || busy) return;
    const request = ++sequence.current;
    const previousRunId = controller.getState().workflowRun?.runId;
    setLocalAction('preview');
    setNotice('');
    try {
      await controller.runWorkflowRecipe(recipe.id, values);
      const completed = controller.getState().workflowRun;
      if (
        request === sequence.current &&
        completed?.status === 'ready' &&
        completed.runId !== previousRunId
      )
        setStaleResultId(undefined);
    } catch (error) {
      if (request === sequence.current) setErrors([{ message: message(error) }]);
    } finally {
      setLocalAction(undefined);
    }
  }

  function save(): void {
    const values = validate();
    if (!values || !recipe) return;
    if (!store) {
      setStorageError(
        'This browser cannot save settings. You can still preview and run workflows.',
      );
      return;
    }
    try {
      store.save({
        schemaVersion: 1,
        recipeId: recipe.id,
        recipeVersion: recipe.version,
        inputs: { ...values, ...(destination.trim() ? { destination: destination.trim() } : {}) },
      });
      setSaved(store.list());
      setStorageError('');
      setNotice('Settings saved on this device. Each run reads the current source data.');
    } catch {
      setStorageError(
        'Settings could not be saved. Check browser storage or clear saved settings and try again.',
      );
    }
  }

  const fieldControl = (field: WorkflowRecipeField): JSX.Element => {
    const fieldId = `${id}-${field.name}`;
    const error = errors.find((item) => item.field === field.name);
    const hint = field.type === 'column' ? undefined : field.description;
    const common = {
      id: fieldId,
      name: field.name,
      disabled: busy,
      'aria-invalid': Boolean(error),
      'aria-describedby': hint || error ? `${fieldId}-hint` : undefined,
    };
    if (field.type === 'boolean')
      return (
        <label className="workflow-check" key={field.name}>
          <input
            {...common}
            type="checkbox"
            aria-describedby={undefined}
            checked={Boolean(inputs[field.name])}
            onChange={(event) => update(field.name, event.target.checked)}
          />
          <span>{field.label}</span>
        </label>
      );
    const value = inputs[field.name];
    const table = field.sourceField
      ? state.analysis?.artifacts.find(
          (item) =>
            item.lineage.operation === 'snapshot' &&
            item.sources.some((source) => source.locator === inputs[field.sourceField!]),
        )
      : undefined;
    return (
      <label
        className={`workflow-field workflow-field-${field.type}`}
        key={field.name}
        htmlFor={fieldId}
      >
        <span>
          {field.label}
          {!field.required && field.default === undefined ? <small>Optional</small> : null}
        </span>
        {field.type === 'column' && table ? (
          <select
            {...common}
            value={value === undefined ? '' : String(value)}
            onChange={(event) =>
              update(field.name, event.target.value === '' ? undefined : Number(event.target.value))
            }
          >
            {!field.required && field.default === undefined && (
              <option value="">Use fixed currency</option>
            )}
            {table.columns.map((column, index) => (
              <option key={column.name} value={index}>
                {index + 1}. {column.label}
              </option>
            ))}
          </select>
        ) : (
          <input
            {...common}
            type={field.type === 'column' ? 'number' : 'text'}
            inputMode={
              field.type === 'decimal' ? 'decimal' : field.type === 'column' ? 'numeric' : undefined
            }
            min={field.type === 'column' ? 1 : undefined}
            max={field.type === 'column' ? 16384 : undefined}
            maxLength={field.type === 'range' ? 1024 : field.type === 'currency' ? 3 : 32}
            list={field.type === 'range' ? `${id}-ranges` : undefined}
            placeholder={
              field.type === 'range'
                ? 'Sheet1!A1:D100'
                : field.type === 'column'
                  ? 'Use fixed currency'
                  : undefined
            }
            value={
              field.type === 'column' && typeof value === 'number' ? value + 1 : String(value ?? '')
            }
            onChange={(event) => {
              const text = event.target.value;
              update(
                field.name,
                field.type === 'column'
                  ? text === ''
                    ? undefined
                    : Number(text) - 1
                  : field.type === 'currency'
                    ? text.toUpperCase()
                    : text,
              );
            }}
          />
        )}
        {(hint || error) && (
          <small id={`${fieldId}-hint`} className={error ? 'workflow-error-copy' : undefined}>
            {error?.message ?? hint}
          </small>
        )}
      </label>
    );
  };

  return (
    <section className="workflow-workbench" aria-labelledby={`${id}-title`}>
      <header className="workflow-heading">
        <div>
          <span className="workflow-eyebrow">Work with your data</span>
          <h2 id={`${id}-title`}>Choose a workflow</h2>
        </div>
        <span className="workflow-badge">{recipes.length} ready to run</span>
      </header>
      <p className="workflow-intro">
        Preview results from your current data, then review where to write them.
      </p>
      <div
        className={`workflow-choices${recipe ? ' workflow-choices-compact' : ''}`}
        role="group"
        aria-label="Data workflows"
      >
        {recipes.map((item) => (
          <button
            key={item.id}
            className="workflow-choice"
            type="button"
            aria-pressed={recipe?.id === item.id}
            disabled={disabled && !localBusy}
            onClick={() => choose(item)}
          >
            <WorkflowIcon id={item.id} />
            <span>
              <strong>{item.title}</strong>
              {!recipe && <small>{item.description}</small>}
            </span>
            {!recipe && (
              <span className="workflow-choice-arrow" aria-hidden="true">
                ↗
              </span>
            )}
          </button>
        ))}
      </div>
      {recipe && (
        <div className="workflow-detail">
          <div className="workflow-step-heading">
            <span aria-hidden="true">1</span>
            <h3>Set up {recipe.title.toLowerCase()}</h3>
          </div>
          {savedRecipe && (
            <div className="workflow-saved">
              <span>Saved settings available</span>
              <button
                type="button"
                disabled={busy}
                onClick={() => {
                  sequence.current += 1;
                  const { destination: savedDestination, ...settings } = savedRecipe.inputs;
                  setInputs(settings);
                  setDestination(typeof savedDestination === 'string' ? savedDestination : '');
                  setErrors([]);
                  setNotice('Saved settings loaded. Preview reads the current source data.');
                  setFocusInputs(true);
                }}
              >
                Use saved settings
              </button>
            </div>
          )}
          <form
            ref={form}
            noValidate
            onSubmit={(event) => {
              event.preventDefault();
              void preview();
            }}
            aria-label={`${recipe.title} settings`}
          >
            <datalist id={`${id}-ranges`}>
              {ranges.map((range) => (
                <option key={range} value={range} />
              ))}
            </datalist>
            <div className="workflow-fields">
              {recipe.fields
                .filter((field) => field.name !== 'destination' && !field.advanced)
                .map(fieldControl)}
            </div>
            <p className="workflow-caption">Column numbers start at 1 within each source range.</p>
            <details className="workflow-advanced">
              <summary>More options</summary>
              <p>{recipe.description}</p>
              <div className="workflow-fields">
                {recipe.fields
                  .filter((field) => field.name !== 'destination' && field.advanced)
                  .map(fieldControl)}
              </div>
            </details>
            {errors.length > 0 && (
              <div className="workflow-error" role="alert" tabIndex={-1} ref={errorPanel}>
                <strong>
                  {errors.some((error) => error.field)
                    ? 'Check these settings'
                    : 'This action could not finish'}
                </strong>
                <ul>
                  {errors.map((error, index) => (
                    <li key={index}>
                      {recipe.fields.find((field) => field.name === error.field)?.label}
                      {error.field ? ': ' : ''}
                      {error.message}
                    </li>
                  ))}
                </ul>
              </div>
            )}
            <div className="workflow-actions">
              <button type="submit" className="workflow-primary" disabled={busy}>
                {localAction === 'preview'
                  ? 'Preparing preview…'
                  : matching && result
                    ? 'Refresh preview'
                    : 'Preview result'}
              </button>
              <button type="button" disabled={busy} onClick={save}>
                Save settings
              </button>
              {localBusy && (
                <button type="button" onClick={() => controller.cancel()}>
                  Cancel
                </button>
              )}
            </div>
            <p className="workflow-caption">
              Preview reads source data. Writing requires your approval.
            </p>
          </form>
          {notice && (
            <p className="workflow-notice" role="status">
              {notice}
            </p>
          )}
          {run?.status === 'running' && matching && (
            <p className="workflow-progress" role="status">
              Reading sources and calculating the result…
            </p>
          )}
          {run?.status === 'cancelled' && matching && (
            <p className="workflow-notice" role="status">
              Preview cancelled. Your settings are ready to try again.
            </p>
          )}
          {run?.status === 'failed' && matching && (
            <div className="workflow-error" role="alert" tabIndex={-1} ref={errorPanel}>
              <strong>Preview could not finish</strong>
              <p>{run.error ?? 'Check your source ranges, then preview again.'}</p>
            </div>
          )}
          {changed && (
            <p className="workflow-notice" role="status">
              Settings changed. Preview again to see the updated result.
            </p>
          )}
          {matching && run?.status === 'ready' && !artifact && (
            <p className="workflow-notice" role="status">
              This preview is no longer available. Preview again to refresh the result.
            </p>
          )}
          {result && (
            <section className="workflow-result" aria-label={`${recipe.title} result`}>
              <div className="workflow-step-heading">
                <span aria-hidden="true">2</span>
                <h3>Review the result</h3>
                <span className="workflow-result-label">
                  {result.truncated
                    ? 'Partial result'
                    : result.rowCount
                      ? 'Preview ready'
                      : 'No rows to write'}
                </span>
              </div>
              <div className="workflow-result-count">
                <strong>{result.rowCount.toLocaleString()}</strong>
                <span>
                  {result.rowCount === 1 ? 'result row' : 'result rows'}
                  <small>{result.columns.length} columns</small>
                </span>
              </div>
              <div className="workflow-sources" aria-label="Result sources">
                {result.sources.map((source) => (
                  <span
                    key={`${source.documentId}:${source.locator}`}
                    title={`Source: ${source.locator}`}
                  >
                    {source.locator}
                  </span>
                ))}
              </div>
              {result.rowCount === 0 ? (
                <p className="workflow-empty">{recipe.result.emptyMessage}</p>
              ) : (
                <>
                  <div
                    className="workflow-table"
                    tabIndex={0}
                    role="region"
                    aria-label={`${recipe.title} result preview`}
                  >
                    <table>
                      <thead>
                        <tr>
                          {result.columns.map((column) => (
                            <th scope="col" key={column.name}>
                              {column.label}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {result.preview.slice(0, 5).map((row, index) => (
                          <tr key={index}>
                            {row.map((cell, column) => (
                              <td key={column}>{String(cell ?? '')}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <p className="workflow-caption">
                    Showing {Math.min(result.preview.length, 5)} of{' '}
                    {result.rowCount.toLocaleString()} rows.
                  </p>
                </>
              )}
              {result.truncated && (
                <p className="workflow-warning" role="status">
                  This result is partial. Use a smaller source range before writing.
                </p>
              )}
              {stale && (
                <p className="workflow-warning" role="alert">
                  A source has changed. Refresh the preview before writing.
                </p>
              )}
              {state.workflowWritesAvailable === false && (
                <p className="workflow-notice" role="status">
                  Writing is unavailable in this document. You can still inspect the result.
                </p>
              )}
              {result.rowCount > 0 &&
                !result.truncated &&
                state.workflowWritesAvailable !== false && (
                  <form
                    className="workflow-write"
                    onSubmit={(event) => {
                      event.preventDefault();
                      if (!canWrite) return;
                      const resultId = result.id;
                      const request = ++sequence.current;
                      setLocalAction('write');
                      void controller
                        .runAnalysis({
                          kind: 'materialize',
                          id: resultId,
                          destination: destination.trim(),
                        })
                        .then(() => {
                          if (request !== sequence.current) return;
                          const failure = controller.getState().error;
                          if (failure?.includes('Source changed:')) setStaleResultId(resultId);
                        })
                        .catch((error: unknown) => {
                          if (request === sequence.current)
                            setErrors([{ message: message(error) }]);
                        })
                        .finally(() => setLocalAction(undefined));
                    }}
                  >
                    <label className="workflow-field" htmlFor={`${id}-destination`}>
                      <span>Write destination</span>
                      <input
                        id={`${id}-destination`}
                        name="destination"
                        placeholder="Results!A1"
                        value={destination}
                        maxLength={1024}
                        disabled={busy}
                        onChange={(event) => setDestination(event.target.value)}
                      />
                    </label>
                    <button className="workflow-primary" type="submit" disabled={!canWrite}>
                      {write?.status === 'pending'
                        ? 'Awaiting your review…'
                        : localAction === 'write'
                          ? 'Checking write…'
                          : 'Review write'}
                    </button>
                    <p className="workflow-caption">
                      Sources are checked again before the write is approved.
                    </p>
                  </form>
                )}
              {write && (
                <p
                  className={`workflow-write-status${write.status === 'uncertain' || write.status === 'failed' ? ' workflow-warning' : ''}`}
                  role="status"
                >
                  {write.status === 'written'
                    ? `Written and verified at ${write.destination}.`
                    : write.status === 'pending'
                      ? `Review the proposed write to ${write.destination}.`
                      : write.status === 'rejected'
                        ? 'Write declined. The preview is still available.'
                        : write.status === 'uncertain'
                          ? 'The write outcome is uncertain. Check Recovery & undo before trying again.'
                          : (write.message ??
                            'The write could not finish. Review the error before trying again.')}
                </p>
              )}
              <details className="workflow-technical">
                <summary>Source versions & technical details</summary>
                <p>
                  Workflow {recipe.id} · version {recipe.version}
                </p>
                <code>{result.id}</code>
                <ul>
                  {result.sources.map((source) => (
                    <li key={`${source.documentId}:${source.locator}`}>
                      {source.locator}
                      <code>{source.hash}</code>
                    </li>
                  ))}
                </ul>
              </details>
            </section>
          )}
        </div>
      )}
      {storageError && (
        <div className="workflow-storage-error" role="status">
          <p>{storageError}</p>
          {store && (
            <button
              type="button"
              disabled={busy}
              onClick={() => {
                try {
                  store.clear();
                  setSaved([]);
                  setStorageError('');
                  setNotice('Saved settings cleared.');
                } catch {
                  setStorageError('Browser storage is unavailable. You can still run workflows.');
                }
              }}
            >
              Clear saved settings
            </button>
          )}
        </div>
      )}
    </section>
  );
}
