import type { ChangeId } from '@ge/contracts';
import type { Proposal } from '../../controller.js';

export interface ProposalCardProps {
  proposals: Proposal[];
  onApply: (changeId: ChangeId) => void;
}

const STATUS_TEXT: Readonly<Record<Proposal['status'], string>> = {
  pending: 'Ready to apply',
  applying: 'Applying…',
  applied: 'Applied as a reversible, tracked change',
  blocked: 'Blocked by a guardrail',
  degraded: 'Anchor drifted — review in panel',
  failed: 'Could not apply',
};

/**
 * The reversible-write review affordance. A staged proposal is shown as a card the user confirms;
 * Accept routes through `applyProposal` → the actuation gate → the bridge, landing a tracked /
 * citation-tagged change carrying agent id, sources, identity, timestamp and a content hash in the
 * host's durable metadata. Nothing is written silently. Mirrors the mockup's "Accept change" card.
 */
export function ProposalCard({ proposals, onApply }: ProposalCardProps): JSX.Element | null {
  if (proposals.length === 0) return null;
  return (
    <section className="proposals" aria-label="Proposed changes">
      {proposals.map((p) => (
        <div
          key={p.changeId}
          className={`card status-${p.status}`}
          role="region"
          aria-label={`Proposed ${p.kind} change: ${p.label}`}
        >
          <div className="card-top" aria-hidden="true" />
          <div className="card-in">
            <div className="cat">Proposed change</div>
            <div className="t">{p.label}</div>
            <div className="w">
              Lands as a reversible, provenanced {p.kind} change — accept or reject it in the host.
            </div>
            <div className={`status-line status-${p.status}`}>{STATUS_TEXT[p.status]}</div>
            {p.detail && <div className="muted small">{p.detail}</div>}
            {p.status === 'pending' && (
              <div className="act">
                <button type="button" className="btn pr" onClick={() => onApply(p.changeId)}>
                  Accept change
                </button>
              </div>
            )}
            {p.status === 'applying' && (
              <div className="act">
                <button type="button" className="btn pr" disabled aria-disabled="true">
                  Applying…
                </button>
              </div>
            )}
          </div>
        </div>
      ))}
    </section>
  );
}
