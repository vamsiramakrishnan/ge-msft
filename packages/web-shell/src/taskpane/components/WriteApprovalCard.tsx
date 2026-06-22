import type { PendingWrite } from '../../controller.js';

export interface WriteApprovalCardProps {
  pending: PendingWrite | undefined;
  onApprove: () => void;
  onReject: () => void;
}

/**
 * The fail-closed, per-write approval affordance for the ADR-0004 command loop. The compiled
 * `ActuationRequest` is rendered **verbatim** as its command line (`set Sales!F2 =C2-D2`) so the
 * user approves exactly what will actuate. The loop is gated on this card: nothing writes until the
 * user Accepts; Reject blocks the write and lets the loop continue with reads. Accessible — the
 * card is a labelled `region`, the command is announced, and both buttons are keyboard-reachable.
 */
export function WriteApprovalCard({
  pending,
  onApprove,
  onReject,
}: WriteApprovalCardProps): JSX.Element | null {
  if (!pending) return null;
  return (
    <section className="card status-pending approval" aria-label="Write approval required">
      <div className="card-top" aria-hidden="true" />
      <div className="card-in">
        <div className="cat">Approve write</div>
        <pre className="cmd" aria-label="Command to approve">
          {pending.command}
        </pre>
        <div className="w">
          This {pending.kind} write will not run until you approve it. Approving lands a reversible,
          provenanced change.
        </div>
        <div className="act">
          <button type="button" className="btn pr" onClick={onApprove}>
            Approve
          </button>
          <button type="button" className="btn" onClick={onReject}>
            Reject
          </button>
        </div>
      </div>
    </section>
  );
}
