import { useState } from 'react';
import type { ArtifactSummary } from '@ge/contracts';
import type { AnalysisAction } from '@ge/runtime';
import type { PanelController, PanelState } from '../../controller.js';

export function AnalysisWorkbench({
  state,
  controller,
  disabled,
}: {
  state: PanelState;
  controller: PanelController;
  disabled: boolean;
}): JSX.Element | null {
  const [range, setRange] = useState('');
  const [headers, setHeaders] = useState(true);
  const [leftId, setLeftId] = useState('');
  const [rightId, setRightId] = useState('');
  const [mapping, setMapping] = useState({
    leftKey: 0,
    leftAmount: 1,
    leftCurrency: 2,
    rightKey: 0,
    rightAmount: 1,
    rightCurrency: 2,
  });
  const [currency, setCurrency] = useState('');
  const [tolerance, setTolerance] = useState('0.01');
  const [destination, setDestination] = useState('');
  const [sql, setSql] = useState('');
  if (!state.analysis) return null;
  const analysis = state.analysis;
  const artifacts = analysis.artifacts;
  const left = artifacts.find((a) => a.id === leftId) ?? artifacts[0];
  const right = artifacts.find((a) => a.id === rightId) ?? artifacts.find((a) => a.id !== left?.id);
  const selected = artifacts.find((a) => a.id === analysis.selected);
  const run = (action: AnalysisAction): void => {
    void controller.runAnalysis(action);
  };
  const columnPicker = (
    table: ArtifactSummary,
    side: 'left' | 'right',
    field: 'Key' | 'Amount' | 'Currency',
  ): JSX.Element => {
    const key = `${side}${field}` as keyof typeof mapping;
    return (
      <label className="analysis-field">
        {field}
        <select
          aria-label={`${side === 'left' ? 'Invoice' : 'Payment'} ${field.toLowerCase()} column`}
          value={mapping[key]}
          onChange={(e) => setMapping((m) => ({ ...m, [key]: Number(e.target.value) }))}
          disabled={disabled || (field === 'Currency' && Boolean(currency))}
        >
          {table.columns.map((c, i) => (
            <option key={c.name} value={i}>
              {c.label}
            </option>
          ))}
        </select>
      </label>
    );
  };
  const mappingValid =
    left &&
    right &&
    left.id !== right.id &&
    ['leftKey', 'leftAmount', ...(currency ? [] : ['leftCurrency'])].every(
      (k) => left.columns[mapping[k as keyof typeof mapping]],
    ) &&
    ['rightKey', 'rightAmount', ...(currency ? [] : ['rightCurrency'])].every(
      (k) => right.columns[mapping[k as keyof typeof mapping]],
    );
  return (
    <details className="analysis-workbench">
      <summary>
        <span>
          <strong>Data workbench</strong>
          <small>Capture · Reconcile · Query</small>
        </span>
        <span className="analysis-count">{artifacts.length} tables</span>
      </summary>
      <div className="analysis-body">
        <p className="analysis-hint">
          Analyze versioned cell snapshots. Review exact results before writing them back.
        </p>
        <form
          onSubmit={(e) => {
            e.preventDefault();
            if (range.trim()) run({ kind: 'capture', range: range.trim(), headers });
          }}
        >
          <label className="analysis-field">
            Source range
            <div className="analysis-inline">
              <input
                aria-label="Source range"
                value={range}
                onChange={(e) => setRange(e.target.value)}
                placeholder="Sheet1!A1:D100"
                disabled={disabled}
                required
              />
              <button type="submit" disabled={disabled || !range.trim()}>
                Capture
              </button>
            </div>
          </label>
          <label className="analysis-check">
            <input
              type="checkbox"
              checked={headers}
              onChange={(e) => setHeaders(e.target.checked)}
              disabled={disabled}
            />
            First row contains headers
          </label>
        </form>
        {artifacts.length > 0 && (
          <div className="analysis-artifacts" aria-label="Analysis tables">
            {artifacts.map((a) => (
              <button
                key={a.id}
                type="button"
                className="analysis-artifact"
                aria-pressed={selected?.id === a.id}
                disabled={disabled}
                onClick={() => run({ kind: 'inspect', id: a.id })}
              >
                <span>{a.title}</span>
                <small>
                  {a.rowCount.toLocaleString()} rows · {a.columns.length} columns
                  {a.truncated ? ' · partial' : ''}
                </small>
              </button>
            ))}
          </div>
        )}
        {artifacts.length >= 2 && (
          <details className="analysis-section">
            <summary>Reconcile invoices and payments</summary>
            {(['left', 'right'] as const).map((side) => {
              const table = side === 'left' ? left : right;
              return table ? (
                <fieldset key={side} disabled={disabled}>
                  <legend>{side === 'left' ? 'Invoices' : 'Payments'}</legend>
                  <label className="analysis-field">
                    Table
                    <select
                      aria-label={`${side === 'left' ? 'Invoice' : 'Payment'} table`}
                      value={table.id}
                      onChange={(e) => {
                        if (side === 'left') setLeftId(e.target.value);
                        else setRightId(e.target.value);
                      }}
                    >
                      {artifacts.map((a) => (
                        <option key={a.id} value={a.id}>
                          {a.title}
                        </option>
                      ))}
                    </select>
                  </label>
                  <div className="analysis-columns">
                    {columnPicker(table, side, 'Key')}
                    {columnPicker(table, side, 'Amount')}
                    {columnPicker(table, side, 'Currency')}
                  </div>
                </fieldset>
              ) : null;
            })}
            <div className="analysis-columns">
              <label className="analysis-field">
                Single currency (optional)
                <input
                  aria-label="Single currency"
                  placeholder="Use currency columns"
                  maxLength={3}
                  value={currency}
                  onChange={(e) => setCurrency(e.target.value.toUpperCase())}
                  disabled={disabled}
                />
              </label>
              <label className="analysis-field">
                Tolerance
                <input
                  aria-label="Reconciliation tolerance"
                  inputMode="decimal"
                  value={tolerance}
                  onChange={(e) => setTolerance(e.target.value)}
                  disabled={disabled}
                />
              </label>
            </div>
            <p className="analysis-hint">
              Amounts are aggregated by key and currency. Invalid values stay visible. Fixed-point
              arithmetic uses six decimal places.
            </p>
            <button
              className="analysis-primary"
              type="button"
              disabled={
                disabled ||
                !mappingValid ||
                !/^\d{1,12}(\.\d{1,6})?$/.test(tolerance) ||
                (Boolean(currency) && !/^[A-Z]{3}$/.test(currency))
              }
              onClick={() => {
                if (left && right)
                  run({
                    kind: 'reconcile',
                    spec: {
                      left: left.id,
                      right: right.id,
                      leftKey: mapping.leftKey,
                      rightKey: mapping.rightKey,
                      leftAmount: mapping.leftAmount,
                      rightAmount: mapping.rightAmount,
                      ...(currency
                        ? { currency }
                        : {
                            leftCurrency: mapping.leftCurrency,
                            rightCurrency: mapping.rightCurrency,
                          }),
                      tolerance,
                    },
                  });
              }}
            >
              Reconcile tables
            </button>
          </details>
        )}
        {selected && (
          <section className="analysis-result" aria-label="Selected analysis result">
            <div className="analysis-result-heading">
              <strong>{selected.title}</strong>
              <button
                type="button"
                className="text-control"
                disabled={disabled}
                onClick={() => run({ kind: 'remove', id: selected.id })}
              >
                Remove
              </button>
            </div>
            <p className="analysis-hint">
              {selected.rowCount.toLocaleString()} rows · showing {selected.preview.length}
              {selected.truncated ? ' · TRUNCATED' : ''}
            </p>
            <div className="analysis-findings" aria-label="Finding actions">
              {analysis.offers.map((offer) => (
                <button
                  type="button"
                  key={offer.id}
                  title={offer.detail}
                  disabled={disabled}
                  onClick={() => run(offer.action)}
                >
                  {offer.title}
                  <span aria-hidden="true"> ↗</span>
                </button>
              ))}
            </div>
            <div
              className="analysis-table-scroll"
              tabIndex={0}
              role="region"
              aria-label={`${selected.title} preview`}
            >
              <table>
                <thead>
                  <tr>
                    {selected.columns.map((c) => (
                      <th key={c.name} scope="col">
                        {c.label}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {selected.preview.map((row, i) => (
                    <tr key={i}>
                      {row.map((cell, j) => (
                        <td key={j}>{String(cell ?? '')}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <details className="analysis-lineage">
              <summary>Source versions & lineage</summary>
              <code>{selected.id}</code>
              <ul>
                {selected.sources.map((s) => (
                  <li key={`${s.documentId}:${s.locator}`}>
                    {s.locator}
                    <small>Version {s.hash.slice(7, 19)}</small>
                  </li>
                ))}
              </ul>
              <p className="analysis-hint">
                {selected.lineage.operation} · {new Date(selected.createdAt).toLocaleTimeString()} ·
                freshness checked before each operation
              </p>
            </details>
            {state.workflowWritesAvailable === false ? (
              <p className="analysis-hint" role="status">
                Writing is unavailable in this document. You can still inspect the result.
              </p>
            ) : (
              <form
                onSubmit={(e) => {
                  e.preventDefault();
                  run({ kind: 'materialize', id: selected.id, destination: destination.trim() });
                }}
              >
                <label className="analysis-field">
                  Write destination
                  <div className="analysis-inline">
                    <input
                      aria-label="Write destination"
                      placeholder="Results!A1"
                      value={destination}
                      onChange={(e) => setDestination(e.target.value)}
                      disabled={disabled}
                      required
                    />
                    <button
                      type="submit"
                      disabled={
                        disabled || !destination.trim() || selected.truncated || !selected.rowCount
                      }
                    >
                      Preview write
                    </button>
                  </div>
                </label>
              </form>
            )}
            {analysis.note && (
              <p role="status" className="analysis-hint">
                {analysis.note}
              </p>
            )}
          </section>
        )}
        {artifacts.length > 0 && (
          <details className="analysis-section">
            <summary>Query with SQL</summary>
            <p className="analysis-hint">
              SELECT queries over the tables listed above. Columns use c0, c1… Input tables are
              restricted to this workspace.
            </p>
            <label className="analysis-field">
              SQL
              <textarea
                aria-label="Analysis SQL"
                rows={5}
                value={sql}
                onChange={(e) => setSql(e.target.value)}
                disabled={disabled}
                placeholder={selected ? `SELECT * FROM ${selected.id}` : 'SELECT …'}
              />
            </label>
            <button
              type="button"
              disabled={disabled || !sql.trim()}
              onClick={() =>
                run({ kind: 'query', inputs: artifacts.map((a) => a.id).slice(0, 16), sql })
              }
            >
              Run query
            </button>
            <div className="analysis-query-ids">
              {artifacts.map((a) => (
                <p key={a.id}>
                  {a.title}
                  <code>{a.id}</code>
                  <small>{a.columns.map((c) => `${c.name}: ${c.label}`).join(' · ')}</small>
                </p>
              ))}
            </div>
          </details>
        )}
        <details className="analysis-section">
          <summary>
            Recovery & undo{' '}
            <span className="analysis-count">{state.recovery?.records.length ?? 0}</span>
          </summary>
          <p className="analysis-hint">
            {state.recovery?.durable
              ? 'Cell-write receipts and previous cells are saved in this document for recovery.'
              : 'Recovery is available for this session only.'}{' '}
            Resume and undo require a fresh preview and approval.
          </p>
          <button type="button" disabled={disabled} onClick={() => run({ kind: 'recovery' })}>
            Refresh recovery
          </button>
          <div className="analysis-recovery">
            {state.recovery?.records.map((r) => (
              <article key={r.id}>
                <strong>{r.target}</strong>
                <span className="analysis-state">{r.state}</span>
                <p className="analysis-hint">
                  {r.rows} × {r.columns} cells · {new Date(r.createdAt).toLocaleString()}
                </p>
                {r.message && <p className="analysis-hint">{r.message}</p>}
                <div className="analysis-inline">
                  {r.canUndo && (
                    <button
                      type="button"
                      disabled={disabled}
                      onClick={() => run({ kind: 'undo', id: r.id })}
                    >
                      Preview undo
                    </button>
                  )}
                  {r.canResume && (
                    <button
                      type="button"
                      disabled={disabled}
                      onClick={() => run({ kind: 'resume', id: r.id })}
                    >
                      Preview resume
                    </button>
                  )}
                  {(r.canForget ||
                    r.canUndo ||
                    r.canResume ||
                    r.state === 'undone' ||
                    r.state === 'superseded') && (
                    <button
                      className="text-control"
                      type="button"
                      disabled={disabled}
                      onClick={() => run({ kind: 'forget', id: r.id })}
                    >
                      Remove receipt
                    </button>
                  )}
                </div>
              </article>
            ))}
          </div>
        </details>
      </div>
    </details>
  );
}
