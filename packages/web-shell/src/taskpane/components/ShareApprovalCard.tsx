import type { PendingShare } from '../../controller.js';

export interface ShareApprovalCardProps {
  pending: PendingShare | undefined;
  onApprove: () => void;
  onReject: () => void;
}

/**
 * The fail-closed approval affordance for `share` — an ADR-0008 **estate**-class write, distinct
 * from an in-document `WriteApprovalCard`: approving this persists `pending.preview`'s full source
 * content to the user's own Microsoft Graph app folder (`/shared`), readable back by every other
 * surface's session, rather than mutating the open Word/Excel/PowerPoint document. The card shows
 * the destination name, what produced the content, and a capped preview so the user approves
 * exactly what leaves the device — never a hidden or partial view of it.
 */
export function ShareApprovalCard({
  pending,
  onApprove,
  onReject,
}: ShareApprovalCardProps): JSX.Element | null {
  if (!pending) return null;
  return (
    <section
      className="card status-pending approval"
      role="region"
      aria-label="Share approval required"
      aria-live="polite"
    >
      <div className="card-top" aria-hidden="true" />
      <div className="card-in">
        <div className="cat">Approve share</div>
        <div className="w">
          Publish <strong>{pending.name}</strong> to your cross-surface handoff store — readable by
          name from any other surface's session — sourced from <em>{pending.sourceLabel}</em>. This
          will not happen until you approve it.
        </div>
        <pre className="cmd" aria-label="Content to be shared, shown verbatim">
          {pending.preview}
          {pending.truncated ? '\n…' : ''}
        </pre>
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
