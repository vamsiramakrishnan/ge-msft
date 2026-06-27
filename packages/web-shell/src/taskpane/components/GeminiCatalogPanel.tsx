import { useEffect, useMemo, useRef, useState } from 'react';
import {
  defaultCatalogSelection,
  type DiscoveryCatalogClient,
  type GeminiCatalog,
  type GeminiCatalogSelection,
  type GeminiCatalogSkill,
} from '@ge/gemini-client';

export interface GeminiCatalogPanelProps {
  catalogClient?: DiscoveryCatalogClient;
  disabled?: boolean;
  onApply: (selection: GeminiCatalogSelection) => void;
}

type LoadState = 'idle' | 'loading' | 'ready' | 'error';

const EMPTY_SKILLS: GeminiCatalogSkill[] = [];
const EMPTY_DATA_STORES: GeminiCatalog['dataStores'] = [];

export function GeminiCatalogPanel({
  catalogClient,
  disabled = false,
  onApply,
}: GeminiCatalogPanelProps): JSX.Element | null {
  const [catalog, setCatalog] = useState<GeminiCatalog | undefined>();
  const [loadState, setLoadState] = useState<LoadState>('idle');
  const [error, setError] = useState<string | undefined>();
  const [expanded, setExpanded] = useState(false);
  const [plannerName, setPlannerName] = useState('');
  const [commandName, setCommandName] = useState('');
  const [defaultName, setDefaultName] = useState('');
  const [dataStoreNames, setDataStoreNames] = useState<Set<string>>(new Set());
  const disabledRef = useRef(disabled);

  useEffect(() => {
    disabledRef.current = disabled;
  }, [disabled]);

  const skills = catalog?.skills ?? EMPTY_SKILLS;
  const dataStores = catalog?.dataStores ?? EMPTY_DATA_STORES;

  const selection = useMemo(
    (): GeminiCatalogSelection => ({
      defaultSkills: defaultName ? findAll(skills, new Set([defaultName])) : [],
      plannerSkill: findOne(skills, plannerName),
      commandSkill: findOne(skills, commandName),
      dataStores: findAll(dataStores, dataStoreNames),
    }),
    [commandName, dataStoreNames, dataStores, defaultName, plannerName, skills],
  );

  const load = (): void => {
    if (!catalogClient || disabledRef.current) return;
    setLoadState('loading');
    setError(undefined);
    const controller = new AbortController();
    catalogClient
      .listCatalog(controller.signal)
      .then((next) => {
        setCatalog(next);
        setError(next.warnings?.join(' '));
        const defaults = defaultCatalogSelection(next);
        setPlannerName(defaults.plannerSkill?.name ?? '');
        setCommandName(defaults.commandSkill?.name ?? '');
        setDefaultName((defaults.defaultSkills ?? [])[0]?.name ?? '');
        setDataStoreNames(new Set((defaults.dataStores ?? []).map((d) => d.name)));
        if (!disabledRef.current) onApply(defaults);
        setLoadState('ready');
      })
      .catch((err: unknown) => {
        setError(err instanceof Error ? err.message : String(err));
        setLoadState('error');
      });
  };

  useEffect(() => {
    load();
    // Load exactly once per client instance; `load` closes over mutable state setters.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [catalogClient]);

  if (!catalogClient) return null;

  return (
    <section
      className="catalog"
      aria-label="Gemini Enterprise catalog"
      aria-disabled={disabled}
      data-expanded={expanded ? 'true' : 'false'}
    >
      <div className="catalog-head">
        <div>
          <div className="eyebrow">Gemini Enterprise catalog</div>
          <div className="catalog-summary">
            {summary(loadState, skills.length, dataStores.length, catalog?.warnings?.length ?? 0)}
          </div>
        </div>
        <div className="catalog-actions">
          <div className="detail-hover">
            <button
              type="button"
              className="icon-btn info"
              aria-label="Catalog details"
              aria-describedby="catalog-details"
            >
              i
            </button>
            <div id="catalog-details" className="detail-popover" role="tooltip">
              Skills are mounted per route: planner for complex plans, command for the constrained
              Office command loop, and default chat only when explicitly selected. Connectors become
              Discovery Engine data stores in the request.
              {error ? <div className="detail-warning">{error}</div> : null}
            </div>
          </div>
          <button
            type="button"
            className="mini-btn"
            onClick={load}
            disabled={disabled || loadState === 'loading'}
          >
            Refresh
          </button>
          <button
            type="button"
            className="mini-btn"
            onClick={() => setExpanded((v) => !v)}
            aria-expanded={expanded}
            disabled={disabled}
          >
            {expanded ? 'Close' : 'Edit'}
          </button>
        </div>
      </div>

      {error ? <div className="catalog-error">Limited catalog</div> : null}

      {catalog && expanded ? (
        <div className="catalog-grid">
          <label className="catalog-field">
            <span>Planner skill</span>
            <select
              value={plannerName}
              disabled={disabled}
              onChange={(event) => setPlannerName(event.target.value)}
            >
              <option value="">None</option>
              {skills.map((skill) => (
                <option key={skill.name} value={skill.name}>
                  {optionLabel(skill)}
                </option>
              ))}
            </select>
          </label>

          <label className="catalog-field">
            <span>Command skill</span>
            <select
              value={commandName}
              disabled={disabled}
              onChange={(event) => setCommandName(event.target.value)}
            >
              <option value="">None</option>
              {skills.map((skill) => (
                <option key={skill.name} value={skill.name}>
                  {optionLabel(skill)}
                </option>
              ))}
            </select>
          </label>

          <label className="catalog-field catalog-wide">
            <span>Default chat skill</span>
            <select
              value={defaultName}
              disabled={disabled}
              onChange={(event) => setDefaultName(event.target.value)}
            >
              <option value="">None</option>
              {skills.map((skill) => (
                <option key={skill.name} value={skill.name}>
                  {optionLabel(skill)}
                </option>
              ))}
            </select>
          </label>

          <fieldset className="catalog-connectors" disabled={disabled}>
            <legend>Connectors</legend>
            {dataStores.length === 0 ? (
              <div className="catalog-empty">No data stores returned for this app.</div>
            ) : (
              dataStores.map((store) => (
                <label key={store.name} className="catalog-check">
                  <input
                    type="checkbox"
                    checked={dataStoreNames.has(store.name)}
                    onChange={(event) => {
                      const next = new Set(dataStoreNames);
                      if (event.target.checked) next.add(store.name);
                      else next.delete(store.name);
                      setDataStoreNames(next);
                    }}
                  />
                  <span>{store.label}</span>
                  {store.suggested ? <span className="catalog-pill">suggested</span> : null}
                </label>
              ))
            )}
          </fieldset>

          <button
            type="button"
            className="catalog-apply"
            disabled={disabled || loadState !== 'ready'}
            onClick={() => onApply(selection)}
          >
            Apply routing
          </button>
        </div>
      ) : null}
    </section>
  );
}

function optionLabel(skill: GeminiCatalogSkill): string {
  return skill.suggestedRoute ? `${skill.label} (${skill.suggestedRoute})` : skill.label;
}

function summary(state: LoadState, skills: number, dataStores: number, warnings: number): string {
  if (state === 'loading') return 'Loading skills and connectors...';
  if (state === 'error') return 'Use fallback routing until the catalog can be listed.';
  if (state === 'ready') {
    const suffix = warnings > 0 ? ' · limited by permissions' : '';
    return `${skills} skills · ${dataStores} connectors${suffix}`;
  }
  return 'Not loaded';
}

function findOne<T extends { name: string }>(items: readonly T[], name: string): T | undefined {
  return name ? items.find((item) => item.name === name) : undefined;
}

function findAll<T extends { name: string }>(items: readonly T[], names: Set<string>): T[] {
  return items.filter((item) => names.has(item.name));
}
