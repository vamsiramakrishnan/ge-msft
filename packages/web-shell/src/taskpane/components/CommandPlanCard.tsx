import type { CommandScope } from '@ge/contracts';
import type { PendingCommandPlan } from '../../controller.js';

export interface CommandPlanCardProps {
  /** The planner's high-level plan awaiting confirm; absent → nothing rendered. */
  pending?: PendingCommandPlan;
  onConfirm: () => void;
  onCancel: () => void;
}

/** A readable label for an orthogonal scope, e.g. "selection" or "range A1:B2". */
function scopeLabel(scope: CommandScope): string {
  return scope.ref ? `${scope.kind} ${scope.ref}` : scope.kind;
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
  onCancel,
}: CommandPlanCardProps): JSX.Element | null {
  if (!pending) return null;
  const { plan } = pending;
  return (
    <section className="command-plan" role="group" aria-label="Proposed plan">
      <div className="command-plan-head">
        <span className="command-plan-verb">/{plan.intent}</span>
        {plan.scope && <span className="command-plan-scope">{scopeLabel(plan.scope)}</span>}
        <span className="command-plan-tag">Plan — confirm before it runs</span>
      </div>
      <ol className="command-plan-steps">
        {plan.steps.map((step, i) => (
          <li key={i}>{step}</li>
        ))}
      </ol>
      {plan.excludes.length > 0 && (
        <ul className="command-plan-excludes" aria-label="Exclusions">
          {plan.excludes.map((ex, i) => (
            <li key={i}>Leave unchanged: {ex}</li>
          ))}
        </ul>
      )}
      {plan.ground.length > 0 && (
        <div className="command-plan-ground">
          Grounded on {plan.ground.map((g) => g.ref ?? g.kind).join(', ')}
        </div>
      )}
      <div className="command-plan-actions">
        <button
          type="button"
          className="command-plan-confirm"
          data-testid="command-plan-confirm"
          onClick={onConfirm}
        >
          Confirm &amp; run
        </button>
        <button
          type="button"
          className="command-plan-cancel"
          data-testid="command-plan-cancel"
          onClick={onCancel}
        >
          Cancel
        </button>
      </div>
    </section>
  );
}
