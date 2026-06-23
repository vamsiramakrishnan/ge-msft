import type { PendingPlan } from '../../controller.js';
import { renderCommandLine } from '../../render-command.js';

export interface PlanApprovalCardProps {
  plan: PendingPlan | undefined;
  onApprove: () => void;
  onReject: () => void;
}

/**
 * The fail-closed, plan-level approval affordance for the ADR-0005 planner/executor. The runtime
 * type-checks and dry-runs the composed turn (reads + pure transforms, no writes), computes the
 * full effect-set, and emits it here for ONE decision. Each effect's `ActuationRequest` is rendered
 * **verbatim** as its command line — and these are the SAME requests that execute on approval, so
 * the user approves exactly what will actuate (no render-benign / execute-malicious divergence).
 *
 * The loop is gated on this card: nothing actuates until the user clicks Approve plan; Reject plan
 * blocks the WHOLE plan. Accessible — the card is a labelled `region`, the effect list is an ordered
 * list announced politely, and both buttons are keyboard-reachable.
 */
export function PlanApprovalCard({
  plan,
  onApprove,
  onReject,
}: PlanApprovalCardProps): JSX.Element | null {
  if (!plan) return null;
  return (
    <section
      className="card status-pending approval plan-approval"
      role="region"
      aria-label="Plan approval required"
      aria-live="polite"
    >
      <div className="card-top" aria-hidden="true" />
      <div className="card-in">
        <div className="cat">Approve plan</div>
        <div className="plan-summary">
          <span className="pin">{plan.summary}</span>
          <span>to review before anything runs</span>
        </div>
        <ol className="plan-effects" aria-label={`Effects in this plan: ${plan.summary}`}>
          {plan.effects.map((effect) => (
            <li key={effect.request.changeId} className="plan-effect">
              <pre className="cmd" aria-label="Effect command, shown verbatim">
                {renderCommandLine(effect.request)}
              </pre>
            </li>
          ))}
        </ol>
        <div className="w">
          Approving runs this whole plan as reversible, provenanced changes — each effect still
          gated and recorded one-by-one. Rejecting blocks the entire plan; nothing runs.
        </div>
        <div className="act">
          <button type="button" className="btn pr" onClick={onApprove}>
            Approve plan
          </button>
          <button type="button" className="btn" onClick={onReject}>
            Reject plan
          </button>
        </div>
      </div>
    </section>
  );
}
