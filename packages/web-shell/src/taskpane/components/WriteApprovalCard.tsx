import type { PendingWrite } from '../../controller.js';
import { kindLabel, uiStrings, type UiLanguage } from '../../ui-language.js';
import { HIGH_IMPACT_KINDS } from './PlanApprovalCard.js';

export interface WriteApprovalCardProps {
  /** The conversation's language, for the card's own text. */
  language?: UiLanguage;
  pending: PendingWrite | undefined;
  onApprove: () => void;
  onReject: () => void;
}

/**
 * The fail-closed, per-write approval affordance for the ADR-0004 command loop. The compiled
 * `ActuationRequest` is rendered **verbatim** as its command line (`set Sales!F2 =C2-D2`) so the
 * user approves exactly what will actuate. The loop is gated on this card: nothing writes until the
 * user presses Apply; Cancel blocks the write and lets the loop continue with reads. Accessible — the
 * card is a labelled `region`, the command is announced, and both buttons are keyboard-reachable.
 */
export function WriteApprovalCard({
  pending,
  onApprove,
  onReject,
  language = 'en',
}: WriteApprovalCardProps): JSX.Element | null {
  if (!pending) return null;
  const t = uiStrings(language);
  return (
    <section
      className="card status-pending approval"
      role="region"
      aria-label="Write approval required"
      aria-live="polite"
    >
      <div className="card-top" aria-hidden="true" />
      <div className="card-in">
        <div className="plan-approval-title">{t.reviewOne}</div>
        <div>
          <span className="plan-effect-kind">{kindLabel(pending.kind, language)}</span>
          {HIGH_IMPACT_KINDS.has(pending.kind) && <span className="effect-review">{t.check}</span>}
        </div>
        <pre className="cmd" aria-label="Command to approve, shown verbatim">
          {pending.command}
        </pre>
        <div className="w">{t.applyNote}</div>
        <div className="act">
          <button type="button" className="btn pr" onClick={onApprove}>
            {t.apply}
          </button>
          <button type="button" className="btn" onClick={onReject}>
            {t.cancel}
          </button>
        </div>
      </div>
    </section>
  );
}
