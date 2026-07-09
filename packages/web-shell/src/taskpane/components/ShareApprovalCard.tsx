import type { PendingShare } from '../../controller.js';

export interface ShareApprovalCardProps {
  pending: PendingShare | undefined;
  onApprove: () => void;
  onReject: () => void;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * The fail-closed approval affordance for `share` — an ADR-0008 **estate**-class write, distinct
 * from an in-document `WriteApprovalCard`: approving this persists content to the user's own
 * Microsoft Graph app folder (`/shared`), readable back by every other surface's session, rather
 * than mutating the open Word/Excel/PowerPoint document. Discloses the ACTUAL total size that will
 * be written (`pending.bytes`) — never just the card's own line-limited preview — so approving
 * never means agreeing to something larger than what's shown.
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
          Publish <strong>{pending.name}</strong> ({formatBytes(pending.bytes)}
          {pending.truncated ? ', truncated to fit the size limit' : ''}) to your cross-surface
          handoff store — readable by name from any other surface's session. This will not happen
          until you approve it.
        </div>
        <div className="w">
          Source: <code>{pending.sourceLabel}</code>
        </div>
        <pre className="cmd" aria-label="Content to be shared, shown verbatim">
          {pending.preview}
          {pending.previewTruncated ? `\n… (${formatBytes(pending.bytes)} total)` : ''}
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
