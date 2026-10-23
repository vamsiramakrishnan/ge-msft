import { useState } from 'react';
import type { PendingPlan, PlanEffect } from '../../controller.js';
import { renderCommandLine } from '../../render-command.js';
import { kindLabel, uiStrings, type UiLanguage } from '../../ui-language.js';

export interface PlanApprovalCardProps {
  /** The conversation's language, for the card's own text. */
  language?: UiLanguage;
  plan: PendingPlan | undefined;
  onRevealTarget?: (target: string) => void;
  onApprove: () => void;
  onReject: () => void;
}

/** The target label for an effect — the dry-run's resolved target, else the request's target. */
function effectTarget(effect: PlanEffect): string | undefined {
  if (effect.dryRun?.target) return effect.dryRun.target;
  const t = effect.request.params.target;
  return t?.range ?? t?.matchText ?? t?.commentId ?? undefined;
}

/** Changes that reach outside the document, delete, or cannot be undone: always flagged for review. */
export const HIGH_IMPACT_KINDS: ReadonlySet<string> = new Set([
  'resolve-revisions',
  'manage-worksheet',
  'delete-slide',
  'find-replace',
  'set-recipients',
  'insert-hyperlink',
  'add-attachment',
]);

/** Any write that carries a file is checked too: it can take document content out with a reply. */
export function isHighImpact(effect: PlanEffect): boolean {
  return (
    HIGH_IMPACT_KINDS.has(effect.request.kind) ||
    effect.request.params.attachment !== undefined ||
    // A table that replaces paragraphs deletes text (review H1).
    effect.request.params.tableGrid?.replaceThrough !== undefined
  );
}

function clipPreview(text: string, max = 60): string {
  const flat = text.replace(/\s+/g, ' ').trim();
  return flat.length > max ? `${flat.slice(0, max)}…` : flat;
}

/**
 * What an effect writes, in plain words, for its collapsed row: the value, the new text, the
 * recipients, the link or the file. The exact command stays in the opened details; this preview is
 * what lets a reviewer approve without opening every row.
 */
export function effectChange(effect: PlanEffect): string | undefined {
  const p = effect.request.params;
  const after = effect.dryRun?.after;
  // Recipients first: a dry-run preview must never hide who the mail goes to.
  // A file first: a reply that carries one must show it even when a body follows.
  const file = p.attachment?.name ? `file: ${p.attachment.name}` : undefined;
  if (p.mail) {
    const parts = (['to', 'cc', 'bcc'] as const)
      .filter((k) => p.mail?.[k]?.length)
      .map((k) => `${k}: ${p.mail?.[k]?.join(', ')}`);
    if (p.mail.subject) parts.push(`subject: ${p.mail.subject}`);
    if (file) parts.unshift(file);
    if (p.mail.body) parts.push(`body: ${p.mail.body}`);
    if (parts.length) return parts.join(' · ');
  }
  if (after !== undefined && after !== '') return `→ ${after}`;
  if (p.hyperlink?.url) return `link: ${p.hyperlink.url}`;
  if (p.attachment) return `attach: ${p.attachment.name ?? p.attachment.uri ?? 'a file'}`;
  if (p.findReplace) return `"${p.findReplace.find}" → "${p.findReplace.replace}"`;
  if (p.worksheet) {
    const ws = p.worksheet;
    return [ws.action, ws.name, ws.newName ? `→ ${ws.newName}` : undefined]
      .filter(Boolean)
      .join(' ');
  }
  if (p.revisions)
    return [p.revisions.action, p.revisions.scope ?? 'revisions'].filter(Boolean).join(' ');
  if (p.text) return `→ ${p.text}`;
  // A new slide: its title, so eight "New slide" rows are told apart before Apply.
  if (p.slide?.title) return p.slide.title;
  const cell = p.cells?.[0]?.[0];
  if (cell !== undefined && cell !== null && String(cell) !== '') return `→ ${String(cell)}`;
  return undefined;
}

/**
 * One reviewable effect in the dry-run effect-set. Collapsed it shows only what changes and where;
 * expanded it reveals the target, the value or before→after preview the no-write dry-run resolved,
 * and the verbatim command line. The command line is the SAME `ActuationRequest` that executes on
 * approval, so what is reviewed is exactly what runs; it is one click away, not shown by default.
 */
function EffectRow({
  effect,
  index,
  language,
  onRevealTarget,
}: {
  effect: PlanEffect;
  index: number;
  language: UiLanguage;
  onRevealTarget?: (target: string) => void;
}): JSX.Element {
  const t = uiStrings(language);
  const [open, setOpen] = useState(false);
  const command = renderCommandLine(effect.request);
  const target = effectTarget(effect);
  const dry = effect.dryRun;
  const change = effectChange(effect);
  const description = target ?? '';
  const highImpact = isHighImpact(effect);
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
        <span className="plan-effect-kind">{kindLabel(effect.request.kind, language)}</span>
        <span className="effect-target-summary">
          {description}
          {highImpact && <span className="effect-review">{t.check}</span>}
          {change && <span className="effect-change">{clipPreview(change)}</span>}
        </span>
        <span className="plan-effect-caret" aria-hidden="true">
          {open ? '▴' : '▾'}
        </span>
      </button>
      {open && (
        <div id={detailsId} className="plan-effect-detail">
          <pre className="cmd" aria-label={`Change ${index} command, shown verbatim`}>
            {command}
          </pre>
          {target && (
            <div className="plan-effect-row">
              <span className="k">{t.where}</span>
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
              <span className="k">{t.becomes}</span>
              <span className="v mono">{dry.resolved}</span>
            </div>
          )}
          {dry?.before !== undefined && dry?.after !== undefined && (
            <div className="plan-effect-row diff" aria-label="Before and after preview">
              <span className="k">{t.beforeAfter}</span>
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
            <div className="muted small">{t.noPreview}</div>
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
 * The loop is gated on this card: nothing actuates until the user clicks Apply; Cancel blocks the
 * WHOLE plan. Accessible — the card is a labelled `region`, the effect list is an ordered
 * list announced politely, each effect head is an expandable button, and both decision buttons are
 * keyboard-reachable. The approve/reject wiring is unchanged and fail-closed.
 */
export function PlanApprovalCard({
  plan,
  onRevealTarget,
  onApprove,
  onReject,
  language = 'en',
}: PlanApprovalCardProps): JSX.Element | null {
  if (!plan) return null;
  const t = uiStrings(language);
  const n = plan.effects.length;
  return (
    <section
      className="card status-pending approval plan-approval"
      role="region"
      aria-label="Plan approval required"
      aria-live="polite"
    >
      <div className="card-top" aria-hidden="true" />
      <div className="card-in">
        <div className="plan-approval-title">{n === 1 ? t.reviewOne : t.reviewMany(n)}</div>
        <ol className="plan-effects" aria-label={n === 1 ? t.reviewOne : t.reviewMany(n)}>
          {plan.effects.map((effect, i) => (
            <EffectRow
              key={effect.request.changeId}
              effect={effect}
              index={i + 1}
              language={language}
              onRevealTarget={onRevealTarget}
            />
          ))}
        </ol>
        <div className="w">{n === 1 ? t.applyNote : t.applyOrder}</div>
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
