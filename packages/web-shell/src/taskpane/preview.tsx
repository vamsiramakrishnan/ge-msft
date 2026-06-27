import { StrictMode, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Surface, ChangeId } from '@ge/contracts';
import type { PanelController, PanelState } from '../controller.js';
import { App } from './components/App.js';
import {
  FIXTURE_CHIPS,
  FIXTURE_MESSAGES,
  FIXTURE_SUGGESTIONS,
  FIXTURE_STEPS,
  FIXTURE_PLAN,
  FIXTURE_PENDING_WRITE,
  FIXTURE_PROPOSALS,
  FIXTURE_SKILLS,
  FIXTURE_ERROR,
  PREVIEW_SURFACES,
} from './preview-fixtures.js';
import './styles.css';
import './preview.css';

/**
 * Standalone preview harness — the "see it" deliverable. Mounts the real <App/> over a fake
 * `PanelController` driven by scripted fixtures, so the whole task pane renders in a plain browser
 * with NO Office host. A toolbar switches the host surface and toggles each card on/off, so every
 * state (streamed message, citations, chips, suggestions, run-steps, pending plan, pending write,
 * proposals, error, busy) is independently inspectable. No network, no controller logic — buttons
 * are wired to harmless console logs so the panel is fully clickable.
 */

/** Which state slices are currently shown — the toolbar drives this into a derived `PanelState`. */
interface Toggles {
  chips: boolean;
  messages: boolean;
  suggestions: boolean;
  skills: boolean;
  steps: boolean;
  plan: boolean;
  write: boolean;
  proposals: boolean;
  error: boolean;
  busy: boolean;
}

const ALL_ON: Toggles = {
  chips: true,
  messages: true,
  suggestions: true,
  skills: true,
  steps: true,
  plan: true,
  write: true,
  proposals: true,
  error: true,
  busy: true,
};

function buildState(t: Toggles): PanelState {
  return {
    messages: t.messages ? FIXTURE_MESSAGES : [],
    chips: t.chips ? FIXTURE_CHIPS : [],
    suggestions: t.suggestions ? FIXTURE_SUGGESTIONS : [],
    proposals: t.proposals ? FIXTURE_PROPOSALS : [],
    changes: [],
    steps: t.steps ? FIXTURE_STEPS : [],
    ...(t.skills ? { skills: FIXTURE_SKILLS } : {}),
    ...(t.write ? { pendingWrite: FIXTURE_PENDING_WRITE } : {}),
    ...(t.plan ? { pendingPlan: FIXTURE_PLAN } : {}),
    busy: t.busy,
    ...(t.error ? { error: FIXTURE_ERROR } : {}),
  };
}

/**
 * A fake `PanelController` exposing the public surface `App` drives. It is a static store: methods
 * log instead of mutating, and `getState`/`subscribe` return the fixture snapshot. Structurally
 * compatible with the bits the view uses; cast to `PanelController` at the call site.
 */
export function makeMockController(state: PanelState): PanelController {
  const log =
    (name: string) =>
    (...args: unknown[]): void => {
      // eslint-disable-next-line no-console
      console.log(`[preview] ${name}`, ...args);
    };
  const mock = {
    getState: () => state,
    subscribe: (_listener: (s: PanelState) => void) => () => undefined,
    refreshContext: async () => log('refreshContext')(),
    attach: async (id: string) => log('attach')(id),
    detach: (id: string) => log('detach')(id),
    send: async (q: string) => log('send')(q),
    runCommands: async (t: string) => log('runCommands')(t),
    cancel: () => log('cancel')(),
    approvePlan: () => log('approvePlan')(),
    rejectPlan: () => log('rejectPlan')(),
    approvePendingWrite: () => log('approvePendingWrite')(),
    rejectPendingWrite: () => log('rejectPendingWrite')(),
    applyProposal: async (id: ChangeId) => log('applyProposal')(id),
    dismissSuggestion: (id: string) => log('dismissSuggestion')(id),
    onAutomate: (q: string) => log('onAutomate')(q),
    registerSkills: (skills: unknown) => log('registerSkills')(skills),
    listSkills: () => state.skills ?? [],
    invokeSkill: async (name: string, args: Record<string, string>) =>
      log('invokeSkill')(name, args),
    onContext: () => undefined,
    onSuggest: () => undefined,
  };
  return mock as unknown as PanelController;
}

const TOGGLE_LABELS: ReadonlyArray<[keyof Toggles, string]> = [
  ['chips', 'Context'],
  ['messages', 'Thread'],
  ['suggestions', 'Suggestions'],
  ['skills', 'Skills'],
  ['steps', 'Run steps'],
  ['plan', 'Plan'],
  ['write', 'Write'],
  ['proposals', 'Proposals'],
  ['error', 'Error'],
  ['busy', 'Busy'],
];

function Preview(): JSX.Element {
  const [surface, setSurface] = useState<Surface>('word');
  const [toggles, setToggles] = useState<Toggles>(ALL_ON);

  // A fresh mock per (surface, toggles) so `getState` reflects the toolbar. `App` re-renders on its
  // own prop change; the mock's `subscribe` is a no-op because the harness is the source of truth.
  const controller = useMemo(() => makeMockController(buildState(toggles)), [toggles]);

  return (
    <div className="preview">
      <aside className="preview-toolbar" aria-label="Preview controls">
        <h1 className="preview-title">Task-pane preview</h1>
        <div className="preview-group">
          <span className="preview-label">Surface</span>
          <div className="preview-surfaces">
            {PREVIEW_SURFACES.map((s) => (
              <button
                key={s}
                type="button"
                className={`preview-chip${s === surface ? ' on' : ''}`}
                onClick={() => setSurface(s)}
                aria-pressed={s === surface}
              >
                {s}
              </button>
            ))}
          </div>
        </div>
        <div className="preview-group">
          <span className="preview-label">Cards</span>
          <div className="preview-toggles">
            {TOGGLE_LABELS.map(([key, label]) => (
              <label key={key} className="preview-toggle">
                <input
                  type="checkbox"
                  checked={toggles[key]}
                  onChange={(e) => setToggles((prev) => ({ ...prev, [key]: e.target.checked }))}
                />
                <span>{label}</span>
              </label>
            ))}
          </div>
        </div>
        <div className="preview-group preview-actions">
          <button type="button" className="preview-btn" onClick={() => setToggles(ALL_ON)}>
            All on
          </button>
          <button
            type="button"
            className="preview-btn"
            onClick={() =>
              setToggles({
                chips: false,
                messages: true,
                suggestions: false,
                skills: false,
                steps: false,
                plan: false,
                write: false,
                proposals: false,
                error: false,
                busy: false,
              })
            }
          >
            Idle / empty
          </button>
        </div>
        <p className="preview-note">
          A fake controller renders the real panel with no Office host. Buttons log to the console.
        </p>
      </aside>
      <div className="preview-stage">
        <div className="preview-frame">
          <App controller={controller} surface={surface} />
        </div>
      </div>
    </div>
  );
}

const el = document.getElementById('root');
if (el) {
  createRoot(el).render(
    <StrictMode>
      <Preview />
    </StrictMode>,
  );
}
