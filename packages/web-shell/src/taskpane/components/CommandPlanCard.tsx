import type { CommandScope } from '@ge/contracts';
import type { PendingCommandPlan } from '../../controller.js';
import { uiStrings, type UiLanguage, type UiStrings } from '../../ui-language.js';

export interface CommandPlanCardProps {
  /** The conversation's language, for the card's own text. */
  language?: UiLanguage;
  /** The planner's high-level plan awaiting confirm; absent → nothing rendered. */
  pending?: PendingCommandPlan;
  onConfirm: () => void;
  /** Ask for a correction: the plan goes back to the planner with what the user types next. */
  onChange?: () => void;
  onCancel: () => void;
}

/**
 * A ref the planner copied from the host, as users should read it. An email ref is an Outlook item
 * id the planner wrote, so it reads "an email" (the card cannot tell it is the open one). Any other
 * single token of 32+ characters is an internal id (a shape or slide id, a content-control id) and
 * reads "an item". Excel sheet names stop at 31 characters, cell ranges carry `!` and links start
 * with http, so readable refs are shown as written.
 */
export function readableRef(ref: string, t: UiStrings): { text: string; own: boolean } | undefined {
  const value = ref.trim();
  if (!value) return undefined;
  if (/^mail:/i.test(value)) return { text: t.planAnEmail, own: true };
  const opaque = /^\S{32,}$/.test(value) && !value.includes('!') && !/^https?:\/\//i.test(value);
  return opaque ? { text: t.planAnItem, own: true } : { text: value, own: false };
}

/**
 * Where the plan works: "In <ref>", "Using your selection", or nothing. The ref is model-written, so
 * it is set apart from the card's own words (see {@link LabelValue}).
 */
function scopeLabel(scope: CommandScope | undefined, t: UiStrings): JSX.Element | undefined {
  const ref = scope?.ref ? readableRef(scope.ref, t) : undefined;
  if (ref)
    return ref.own ? (
      <>{`${t.planIn} ${ref.text}`}</>
    ) : (
      <LabelValue label={t.planIn} value={ref.text} />
    );
  return scope?.kind === 'selection' ? <>{t.planSelection}</> : undefined;
}

/**
 * A card label with a model-written value. The value sits in its own element with a distinct look,
 * so planner text ("everything (read-only)") cannot pass for the card's own wording.
 */
function LabelValue({ label, value }: { label: string; value: string }): JSX.Element {
  return (
    <>
      {label} <span className="model-value">{value}</span>
    </>
  );
}

/**
 * The planner-confirm card (EXPERIENCE.md §F). For a complex free-text actuating request, the planner
 * proposes a high-level {@link PendingCommandPlan} — the verb, the scope, the ordered steps, and any
 * exclusions/grounding — which the user confirms (or cancels) BEFORE the executor runs. This is the
 * legible "see the plan before it acts" stage; on confirm the executor then stages its own
 * effect-level gate (`PlanApprovalCard`). Nothing actuates from this card directly.
 */
export function CommandPlanCard({
  pending,
  onConfirm,
  onChange,
  onCancel,
  language = 'en',
}: CommandPlanCardProps): JSX.Element | null {
  if (!pending) return null;
  const { plan } = pending;
  const t = uiStrings(language);
  const where = scopeLabel(plan.scope, t);
  // Named sources only: bare keywords ("this", "unit") are planner vocabulary, not words for users.
  const refs = plan.ground.flatMap((g) => (g.ref ? [readableRef(g.ref, t)] : []));
  const ownWords = [...new Set(refs.flatMap((r) => (r?.own ? [r.text] : [])))];
  const sources = refs.flatMap((r) => (r && !r.own ? [r.text] : []));
  return (
    <section className="command-plan" role="group" aria-label="Proposed plan">
      <div className="command-plan-head">
        <span className="command-plan-title">{t.planTitle}</span>
        {where && <span className="command-plan-scope">{where}</span>}
      </div>
      <ol className="command-plan-steps">
        {plan.steps.map((step, i) => (
          <li key={i}>{step}</li>
        ))}
      </ol>
      {plan.excludes.length > 0 && (
        <ul className="command-plan-excludes" aria-label="Exclusions">
          {plan.excludes.map((ex, i) => (
            <li key={i}>
              <LabelValue label={t.planKeep} value={ex} />
            </li>
          ))}
        </ul>
      )}
      {ownWords.length + sources.length > 0 && (
        <div className="command-plan-ground">
          {sources.length > 0 ? (
            <LabelValue
              label={[t.planUsing, ...ownWords].join(' ') + (ownWords.length ? ',' : '')}
              value={sources.join(', ')}
            />
          ) : (
            `${t.planUsing} ${ownWords.join(', ')}`
          )}
        </div>
      )}
      <div className="command-plan-actions">
        <button
          type="button"
          className="btn pr"
          data-testid="command-plan-confirm"
          onClick={onConfirm}
        >
          {t.planGo}
        </button>
        {onChange ? (
          <button
            type="button"
            className="btn"
            data-testid="command-plan-change"
            onClick={onChange}
          >
            {t.planChange}
          </button>
        ) : null}
        <button
          type="button"
          className="btn danger"
          data-testid="command-plan-cancel"
          onClick={onCancel}
        >
          {t.cancel}
        </button>
      </div>
    </section>
  );
}
