import { useState } from 'react';
import type { ChangeId } from '@ge/contracts';
import type { Proposal } from '../../controller.js';
import { ProvenanceDetail } from './ProvenanceDetail.js';

export interface ProposalCardProps {
  proposals: Proposal[];
  disabled?: boolean;
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

/** A short verb noun for the proposal kind, used in the "Lands as…" copy. */
const KIND_NOUN: Readonly<Record<string, string>> = {
  'tracked-change': 'tracked-change (redline)',
  'write-cells': 'formula',
  'add-comment': 'comment',
  'comment-reply': 'comment reply',
  'format-cells': 'format',
};

function firstCell(cells: readonly (readonly string[])[] | undefined): string | undefined {
  return cells?.[0]?.[0];
}

/**
 * The surface-faithful body of a proposal: an Excel `write-cells` shows the value formula-first as
 * `=…` against its range target (mockup `2-excel.html`'s formula bar); a Word `tracked-change`
 * shows the redline as old→new struck/inserted text (mockup `1-word.html`'s wavy redline). Other
 * kinds fall back to the proposal label. Presentational only — the actuation is unchanged.
 */
function ProposalBody({ proposal }: { proposal: Proposal }): JSX.Element {
  const p = proposal.params;
  if (proposal.kind === 'write-cells') {
    const value = firstCell(p.cells) ?? p.text ?? '';
    const range = p.target?.range;
    return (
      <div className="proposal-formula" aria-label="Cell write, formula-first">
        {range && <span className="cell-target mono">{range}</span>}
        <code className="formula">{value}</code>
      </div>
    );
  }
  if (proposal.kind === 'tracked-change') {
    const oldText = p.target?.matchText;
    const newText = p.text ?? p.html ?? p.ooxml ?? '';
    return (
      <div className="proposal-redline" aria-label="Tracked change, redline">
        {oldText && <del className="redline-del">{oldText}</del>}
        <ins className="redline-ins">{newText}</ins>
      </div>
    );
  }
  return <div className="t">{proposal.label}</div>;
}

/** The Excel linked-entity card (mockup `2-excel.html`'s `◆ <entity>` expandable card). */
function EntityCardView({ card }: { card: NonNullable<Proposal['entityCard']> }): JSX.Element {
  return (
    <div className="entity-card" role="group" aria-label={`Entity: ${card.title}`}>
      <div className="entity-head">
        <span className="entity-mark" aria-hidden="true">
          ◆
        </span>
        <span className="entity-title">{card.title}</span>
        {card.subtitle && <span className="entity-sub">{card.subtitle}</span>}
      </div>
      <dl className="entity-rows">
        {card.rows.map((r) => (
          <div key={r.key} className="entity-row">
            <dt>{r.key}</dt>
            <dd>{r.value}</dd>
          </div>
        ))}
      </dl>
      {card.footnote && <div className="entity-foot muted small">{card.footnote}</div>}
    </div>
  );
}

function ProposalItem({
  proposal,
  disabled,
  onApply,
}: {
  proposal: Proposal;
  disabled: boolean;
  onApply: (changeId: ChangeId) => void;
}): JSX.Element {
  const [showProvenance, setShowProvenance] = useState(false);
  const kindNoun = KIND_NOUN[proposal.kind] ?? `${proposal.kind} change`;
  const detailId = `provenance-${proposal.changeId}`;
  return (
    <div
      className={`card status-${proposal.status}`}
      role="region"
      aria-label={`Proposed ${proposal.kind} change: ${proposal.label}`}
    >
      <div className="card-top" aria-hidden="true" />
      <div className="card-in">
        <div className="cat">Proposed change</div>
        <div className="t">{proposal.label}</div>
        <ProposalBody proposal={proposal} />
        {proposal.entityCard && <EntityCardView card={proposal.entityCard} />}
        <div className="w">
          Lands as a reversible, provenanced {kindNoun} — accept or reject it in the host.
        </div>
        <div className={`status-line status-${proposal.status}`}>
          {STATUS_TEXT[proposal.status]}
        </div>
        {proposal.detail && <div className="muted small">{proposal.detail}</div>}
        {proposal.provenance && (
          <div className="provenance-drill">
            <button
              type="button"
              className="link prov-toggle"
              aria-expanded={showProvenance}
              aria-controls={detailId}
              onClick={() => setShowProvenance((v) => !v)}
            >
              {showProvenance ? 'Hide provenance' : 'Show provenance'}
            </button>
            {showProvenance && (
              <div id={detailId}>
                <ProvenanceDetail provenance={proposal.provenance} />
              </div>
            )}
          </div>
        )}
        {proposal.status === 'pending' && (
          <div className="act">
            <button
              type="button"
              className="btn pr"
              disabled={disabled}
              onClick={() => onApply(proposal.changeId)}
            >
              Accept change
            </button>
          </div>
        )}
        {proposal.status === 'applying' && (
          <div className="act">
            <button type="button" className="btn pr" disabled aria-disabled="true">
              Applying…
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

/**
 * The reversible-write review affordance. A staged proposal is shown as a card the user confirms;
 * Accept routes through `applyProposal` → the actuation gate → the bridge, landing a tracked /
 * citation-tagged change carrying agent id, sources, identity, timestamp and a content hash in the
 * host's durable metadata. Nothing is written silently. The body is rendered surface-faithfully
 * (Excel formula-first, Word redline) and an applied write can drill into its provenance.
 */
export function ProposalCard({
  proposals,
  onApply,
  disabled = false,
}: ProposalCardProps): JSX.Element | null {
  if (proposals.length === 0) return null;
  return (
    <section className="proposals" aria-label="Proposed changes">
      {proposals.map((p) => (
        <ProposalItem disabled={disabled} key={p.changeId} proposal={p} onApply={onApply} />
      ))}
    </section>
  );
}
