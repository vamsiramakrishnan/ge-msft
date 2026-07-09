import { useState } from 'react';
import type { ActuationRequest } from '@ge/contracts';
import type { PendingPlan, PlanEffect } from '../../controller.js';
import { renderCommandLine } from '../../render-command.js';

export interface PlanApprovalCardProps {
  plan: PendingPlan | undefined;
  onRevealTarget?: (target: string) => void;
  onApprove: () => void;
  onReject: () => void;
}

/** A short human noun for an effect kind, used as the effect-row eyebrow. */
function effectKindLabel(kind: ActuationRequest['kind']): string {
  switch (kind) {
    case 'write-cells':
      return 'write';
    case 'tracked-change':
      return 'suggest';
    case 'add-comment':
    case 'comment-reply':
      return 'comment';
    case 'format-cells':
      return 'format';
    default:
      return kind;
  }
}

/** The target label for an effect — the dry-run's resolved target, else the request's target. */
function effectTarget(effect: PlanEffect): string | undefined {
  if (effect.dryRun?.target) return effect.dryRun.target;
  const t = effect.request.params.target;
  return t?.range ?? t?.matchText ?? t?.commentId ?? undefined;
}

/**
 * One reviewable effect in the dry-run effect-set. Collapsed it shows the verbatim command line;
 * expanded it reveals the target and — when the runtime's no-write dry-run resolved it — the value
 * the effect will produce and/or a before→after preview. The command line is the SAME
 * `ActuationRequest` that executes on approval, so what is reviewed is exactly what runs.
 */
function EffectRow({
  effect,
  index,
  onRevealTarget,
}: {
  effect: PlanEffect;
  index: number;
  onRevealTarget?: (target: string) => void;
}): JSX.Element {
  const [open, setOpen] = useState(false);
  const command = renderCommandLine(effect.request);
  const target = effectTarget(effect);
  const dry = effect.dryRun;
  const detailsId = `plan-effect-${effect.request.changeId}`;
  return (
    <li className="plan-effect">
      <button
        type="button"
        className="plan-effect-head"
        aria-expanded={open}
        aria-controls={detailsId}
        onClick={() => setOpen((v) => !v)}
      >
        <span className="plan-effect-kind eyebrow">{effectKindLabel(effect.request.kind)}</span>
        <pre className="cmd" aria-label={`Effect ${index} command, shown verbatim`}>
          {command}
        </pre>
        <span className="plan-effect-caret" aria-hidden="true">
          {open ? '-' : '+'}
        </span>
      </button>
      {open && (
        <div id={detailsId} className="plan-effect-detail">
          {target && (
            <div className="plan-effect-row">
              <span className="k">Target</span>
              {onRevealTarget ? (
                <button
                  type="button"
                  className="v mono host-target-link"
                  onClick={() => onRevealTarget(target)}
                  title="Open this target in the host"
                >
                  {target}
                </button>
              ) : (
                <span className="v mono">{target}</span>
              )}
            </div>
          )}
          {dry?.resolved !== undefined && (
            <div className="plan-effect-row">
              <span className="k">Resolves to</span>
              <span className="v mono">{dry.resolved}</span>
            </div>
          )}
          {dry?.before !== undefined && dry?.after !== undefined && (
            <div className="plan-effect-row diff" aria-label="Before and after preview">
              <span className="k">Change</span>
              <span className="v">
                <span className="diff-before">{dry.before}</span>
                <span className="diff-arrow" aria-hidden="true">
                  →
                </span>
                <span className="diff-after">{dry.after}</span>
              </span>
            </div>
          )}
          {!target && dry?.resolved === undefined && dry?.before === undefined && (
            <div className="muted small">
              The dry-run could not resolve a preview for this effect — it runs as shown above.
            </div>
          )}
        </div>
      )}
    </li>
  );
}

/**
 * The fail-closed, plan-level approval affordance for the ADR-0005 planner/executor. The runtime
 * type-checks and dry-runs the composed turn (reads + pure transforms, no writes), computes the
 * full effect-set, and emits it here for ONE decision. Each effect's `ActuationRequest` is rendered
 * **verbatim** as its command line — and these are the SAME requests that execute on approval, so
 * the user approves exactly what will actuate (no render-benign / execute-malicious divergence).
 * Each effect expands to its target and the dry-run's resolved value / before→after preview: a
 * reviewable program before it runs.
 *
 * The loop is gated on this card: nothing actuates until the user clicks Approve plan; Reject plan
 * blocks the WHOLE plan. Accessible — the card is a labelled `region`, the effect list is an ordered
 * list announced politely, each effect head is an expandable button, and both decision buttons are
 * keyboard-reachable. The approve/reject wiring is unchanged and fail-closed.
 */
export function PlanApprovalCard({
  plan,
  onRevealTarget,
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
          {plan.effects.map((effect, i) => (
            <EffectRow
              key={effect.request.changeId}
              effect={effect}
              index={i + 1}
              onRevealTarget={onRevealTarget}
            />
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
