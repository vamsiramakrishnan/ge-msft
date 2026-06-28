import { useEffect, useMemo, useRef, useState } from 'react';
import {
  deriveOutput,
  actionParameters,
  hasUnfilledPlaceholder,
  GROUND_SOURCE_TO_SELECTION_KIND,
  type Surface,
  type ChangeId,
  type QuickAction,
  type Intent,
  type GroundingSelection,
} from '@ge/contracts';
import {
  applyCatalogSelection,
  resolveGrounding,
  type DiscoveryCatalogClient,
  type GeminiCatalogSelection,
  type GroundingResolveContext,
  type ResolvedGrounding,
} from '@ge/gemini-client';
import type { PanelController } from '../../controller.js';
import { usePanelState } from '../usePanelState.js';
import { Toolbar } from './Toolbar.js';
import { MessageThread } from './MessageThread.js';
import { Composer, type ComposerInvocation, type ComposerMention } from './Composer.js';
import { invocationToSeed, quickActionToInvocation } from './quick-action-seed.js';
import { QuickActionParamForm } from './QuickActionParamForm.js';
import { ProposalCard } from './ProposalCard.js';
import { RunSteps } from './RunSteps.js';
import { WriteApprovalCard } from './WriteApprovalCard.js';
import { PlanApprovalCard } from './PlanApprovalCard.js';
import { CommandPlanCard } from './CommandPlanCard.js';
import { GeminiCatalogPanel } from './GeminiCatalogPanel.js';
import { surfacePrimaryActions } from './SurfaceCommandCenter.js';

export interface AppProps {
  controller: PanelController;
  surface: Surface;
  agentLabel?: string;
  /**
   * The intents the active surface can actually run (from `intentsForManifest(bridge manifest)`).
   * Narrows the quick-action bar and the `/` palette to capability closure (ADR-0006). When omitted
   * (preview/tests), only the per-surface filter applies.
   */
  allowedIntents?: Iterable<Intent>;
  catalogClient?: DiscoveryCatalogClient;
  onCatalogRouting?: (selection: ReturnType<typeof applyCatalogSelection>) => void;
}

const SURFACE_PLACEHOLDER: Readonly<Record<string, string>> = {
  word: 'Ask about the selection…',
  excel: 'Ask about this range…',
  powerpoint: 'Ask about this slide…',
  onenote: 'Ask about this page…',
  outlook: 'Ask about this email…',
  teams: 'Ask about this meeting…',
};

/**
 * THE single routing predicate (EXPERIENCE.md §3) — total over the 7-verb set. Chat verbs
 * (`ask`/`summarize`/`explain`) and a bare question (undefined intent) are single-shot reads → `send`;
 * the four specialist verbs (`rewrite`/`review`/`draft`/`notes`) land an effect through the gate →
 * `runCommands`. A `rewrite` can NEVER reach `send`. App makes this decision ONCE, here, used by both
 * the quick-action chips and the composer submit.
 */
export function isActuating(intent: Intent | undefined): boolean {
  // Single source of truth (no second hand-maintained verb set to drift): a verb actuates iff its
  // closure-derived output is not `chat`. `deriveOutput` lives in contracts next to INTENT_REQUIRES,
  // so the routing decision and the capability closure can never disagree (security review, Finding 2).
  return intent !== undefined && deriveOutput(intent) !== 'chat';
}

/** Words/markers that signal a multi-step or constrained instruction worth a plan-confirm first. */
const CONSTRAINT_MARKERS =
  /(^|\s)(but|only|except|without|unless|then|leave|keep|preserve|ignore|excluding)(\s|$)|[,;]/i;

/**
 * The heuristic for "complex free-text" (EXPERIENCE.md §F): an instruction earns the planner-confirm
 * front door when it's long (≥ 12 words) or carries a constraint/exclusion marker (",", "but",
 * "only", "leave …", etc.). A short, single-shot instruction ("make it formal") skips the planner and
 * goes straight to the executor. Conservative by design — the executor's gate is the backstop either way.
 */
export function isComplexInstruction(instruction: string): boolean {
  const t = instruction.trim();
  if (!t) return false;
  return t.split(/\s+/).length >= 12 || CONSTRAINT_MARKERS.test(t);
}

const EXCEL_CHART_CREATE_RE =
  /^\s*(?:please\s+)?(?:(?:can|could|would)\s+you\s+)?(?:create|make|insert|add|build|generate|plot|visuali[sz]e)\b[\s\S]*\b(?:chart|graph|visuali[sz]ation)\b/i;
const EXCEL_CHART_CONVERT_RE =
  /^\s*(?:please\s+)?(?:turn|convert)\b[\s\S]*\b(?:into|to)\b[\s\S]*\b(?:chart|graph|visuali[sz]ation)\b/i;
const WORD_REWRITE_RE =
  /^\s*(?:please\s+)?(?:(?:can|could|would)\s+you\s+)?(?:rewrite|revise|tighten|edit|replace|improve)\b[\s\S]*\b(?:selection|selected text|paragraph|text|wording|clause|sentence)\b/i;
const WORD_REVIEW_RE =
  /^\s*(?:please\s+)?(?:(?:can|could|would)\s+you\s+)?(?:review|comment|flag|mark)\b[\s\S]*\b(?:issue|issues|risk|risks|gap|gaps|claim|claims|comment|comments)\b/i;
const POWERPOINT_DRAFT_RE =
  /^\s*(?:please\s+)?(?:(?:can|could|would)\s+you\s+)?(?:create|make|insert|add|build|generate|draft)\b[\s\S]*\bslides?\b/i;
const OUTLOOK_DRAFT_RE =
  /^\s*(?:please\s+)?(?:(?:can|could|would)\s+you\s+)?(?:draft|write|compose|create)\b[\s\S]*\b(?:reply|email|mail|message)\b/i;

function intentAllowed(intent: Intent, allowedIntents: Iterable<Intent> | undefined): boolean {
  if (allowedIntents === undefined) return true;
  for (const allowed of allowedIntents) {
    if (allowed === intent) return true;
  }
  return false;
}

/**
 * Narrow natural-language promotion for bridge-backed mutations. Plain "create a chart..." / "draft
 * a slide..." requests are actuating requests, but historically landed in read-only chat unless the
 * user knew to type a slash verb. Promote only imperative language, only when the runtime capability
 * closure includes the matching intent. The promoted turn still enters `runCommands`, so it dry-runs,
 * previews, and waits for explicit approval before any host mutation.
 */
export function inferImplicitIntent(
  surface: Surface,
  allowedIntents: Iterable<Intent> | undefined,
  inv: ComposerInvocation,
): Intent | undefined {
  if (inv.intent !== undefined) return inv.intent;
  const raw = inv.raw.trim();
  if (!raw || raw.startsWith('/')) return undefined;
  switch (surface) {
    case 'excel':
      return intentAllowed('visualize', allowedIntents) &&
        (EXCEL_CHART_CREATE_RE.test(raw) || EXCEL_CHART_CONVERT_RE.test(raw))
        ? 'visualize'
        : undefined;
    case 'word':
      if (intentAllowed('rewrite', allowedIntents) && WORD_REWRITE_RE.test(raw)) return 'rewrite';
      if (intentAllowed('review', allowedIntents) && WORD_REVIEW_RE.test(raw)) return 'review';
      return undefined;
    case 'powerpoint':
      return intentAllowed('draft', allowedIntents) && POWERPOINT_DRAFT_RE.test(raw)
        ? 'draft'
        : undefined;
    case 'outlook':
      return intentAllowed('draft', allowedIntents) && OUTLOOK_DRAFT_RE.test(raw)
        ? 'draft'
        : undefined;
    case 'onenote':
    case 'teams':
      return undefined;
  }
}

/**
 * Map ONE typed composer `@`-mention to its {@link GroundingSelection} (Finding #2/#B-wire). The
 * `GroundSource` kind picks the value-level selection kind; the reference-kinds (`current-context`,
 * `unit`) carry no id, while `document`/`person`/`data-store`/`upload` need the addressable handle the
 * mention carried (`ref`). A mention of an addressable kind with NO ref cannot be resolved — it is
 * dropped (returns `undefined`) rather than smuggled into the prompt as raw text.
 */
export function mentionToSelection(m: ComposerMention): GroundingSelection | undefined {
  const kind = GROUND_SOURCE_TO_SELECTION_KIND[m.kind];
  switch (kind) {
    case 'current-context':
    case 'unit':
      return { kind };
    case 'document':
    case 'person':
    case 'data-store':
      return m.ref ? { kind, id: m.ref } : undefined;
    case 'upload':
      return m.ref ? { kind, fileId: m.ref } : undefined;
    default:
      return undefined;
  }
}

/**
 * Turn an invocation's typed `@`-mentions into STRUCTURED grounding (Finding #2/#B-wire): map each
 * mention to a {@link GroundingSelection}, then resolve them onto the streamAssist request buckets via
 * `resolveGrounding` (gemini-client). The mentions become addressed `queryParts`/`dataStoreSpecs`/
 * `fileIds` — NEVER inlined into the prompt string; the raw text is kept only for audit/display.
 * Returns `undefined` when no mention resolved to anything, so a mention-free turn carries no grounding.
 */
export function invocationToGrounding(
  inv: ComposerInvocation,
  ctx: GroundingResolveContext = {},
): ResolvedGrounding | undefined {
  const selections = inv.mentions
    .map(mentionToSelection)
    .filter((s): s is GroundingSelection => s !== undefined);
  if (selections.length === 0) return undefined;
  return resolveGrounding(selections, ctx);
}

/**
 * The task pane. A thin React view over `PanelController` state: header (agent identity), context
 * tray (attach/detach chips), streamed grounded thread with citations, proposal-review cards, and
 * the composer (send / cancel). No host or network code here — the controller owns all of that.
 */
export function App({
  controller,
  surface,
  agentLabel,
  allowedIntents,
  catalogClient,
  onCatalogRouting,
}: AppProps): JSX.Element {
  const state = usePanelState(controller);
  // The parameterized action awaiting its `{{name}}` fill values (Workstream H), or undefined.
  const [paramFill, setParamFill] = useState<QuickAction | undefined>(undefined);
  // Catalog routing is admin-grade config; it lives behind this settings toggle, not as always-on
  // chrome, so the default pane stays task-focused (quiet by default). The catalog stays mounted
  // either way so its default routing still loads on open.
  const [settingsOpen, setSettingsOpen] = useState(false);
  // Settings (catalog routing) is config, so it earns a real modal dialog rather than inline chrome —
  // it overlays without costing the default pane any vertical space. Driven imperatively off
  // `settingsOpen`; guarded so jsdom / hosts without `showModal` degrade rather than throw.
  const settingsDialogRef = useRef<HTMLDialogElement>(null);
  useEffect(() => {
    const dlg = settingsDialogRef.current;
    if (!dlg) return;
    if (settingsOpen && typeof dlg.showModal === 'function' && !dlg.open) {
      try {
        dlg.showModal();
      } catch {
        /* unsupported host (e.g. jsdom) — the panel still renders, just not as a top-layer modal */
      }
    } else if (!settingsOpen && dlg.open) {
      dlg.close();
    }
  }, [settingsOpen]);
  const hasBlockingGate = Boolean(
    state.pendingCommandPlan ?? state.pendingPlan ?? state.pendingWrite,
  );
  const actionBlocked = state.busy || hasBlockingGate;
  const primaryActions = useMemo(
    () => surfacePrimaryActions(surface, allowedIntents),
    [allowedIntents, surface],
  );
  const primaryActionIds = useMemo(
    () => primaryActions.map((action) => action.id),
    [primaryActions],
  );
  const attachedCount = state.chips.filter((chip) => chip.attached).length;
  const availableCount = Math.max(0, state.chips.length - attachedCount);
  const proposalCount = state.proposals.filter((proposal) => proposal.status === 'pending').length;

  // Load the attachable-context chips once on mount.
  useEffect(() => {
    void controller.refreshContext();
  }, [controller]);

  const onToggle = (id: string, attach: boolean): void => {
    if (attach) void controller.attach(id);
    else controller.detach(id);
  };

  // Dispatch the typed invocation through ONE predicate-routed seam (EXPERIENCE.md §3): chat verbs
  // route to `send`, write/annotation verbs to the fail-closed plan/approval loop (`runCommands`).
  // The model-facing TASK stays the CLI seed (deterministic from the typed fields); ALONGSIDE it, the
  // typed `@`-mentions are CONSUMED into STRUCTURED grounding (Finding #2/#B-wire) via
  // `invocationToGrounding` → `resolveGrounding`, and passed as the turn's grounding — never discarded
  // nor forwarded only as raw text. No new gate is introduced; grounding only scopes the existing route.
  const dispatch = (inv: ComposerInvocation): void => {
    const effectiveIntent = inferImplicitIntent(surface, allowedIntents, inv);
    const routedInv =
      effectiveIntent !== undefined && effectiveIntent !== inv.intent
        ? { ...inv, intent: effectiveIntent, raw: `/${effectiveIntent} ${inv.raw}`.trim() }
        : inv;
    const seed = invocationToSeed(routedInv);
    // Fail-closed (Workstream H): a typed `{{name}}` slot must be filled before dispatch — never send
    // a literal placeholder to the model. A parameterized chip is collected via QuickActionParamForm
    // first, so this guard only fires on a defective seed; drop it rather than actuate on raw braces.
    if (hasUnfilledPlaceholder(seed)) return;
    const grounding = invocationToGrounding(inv);
    // EXPERIENCE.md §F — the planner-confirm front door, enforced ONLY for complex free-text: a
    // composer-typed (raw ≠ '') actuating instruction with constraints first proposes a confirmable
    // CommandPlan. A chip/preset (raw === '') or a simple instruction routes straight to the executor.
    const composerOrigin = inv.raw.trim() !== '';
    if (composerOrigin && isActuating(effectiveIntent) && isComplexInstruction(inv.instruction)) {
      void controller.proposePlan(seed, grounding);
    } else if (isActuating(effectiveIntent)) {
      void controller.runCommands(seed, grounding);
    } else {
      void controller.send(seed, grounding);
    }
  };

  // A chip is the same typed Invocation a composer line is — pre-filled, then dispatched through the
  // one shared `dispatch`. (`action.output` and `isActuating` agree by construction, ADR-0006.) A
  // parameterized action (Workstream H) opens the fill form FIRST instead of dispatching, so its
  // `{{name}}` slots are collected before the typed invocation is built.
  const onQuickAction = (action: QuickAction): void => {
    if (actionParameters(action).length > 0) setParamFill(action);
    else dispatch(quickActionToInvocation(action));
  };

  // Dispatch a parameterized action once its fill values are collected: substitute the values into
  // the prompt, build the SAME typed invocation a bare chip would (scope/intent/grounding intact),
  // and route it through the one shared `dispatch`.
  const onParamFillSubmit = (action: QuickAction, values: Record<string, string>): void => {
    setParamFill(undefined);
    dispatch(quickActionToInvocation(action, values));
  };

  // A structured composer submit (`/verb` intent + scope + `@`-mentions + instruction) → the same path.
  const onInvoke = (inv: ComposerInvocation): void => dispatch(inv);

  return (
    <div className="panel" data-surface={surface} aria-busy={state.busy}>
      <Toolbar
        surface={surface}
        allowedIntents={allowedIntents}
        agentLabel={agentLabel}
        busy={actionBlocked}
        hasGate={hasBlockingGate}
        chips={state.chips}
        attachedCount={attachedCount}
        availableCount={availableCount}
        messageCount={state.messages.length}
        proposalCount={proposalCount}
        skills={state.skills ?? []}
        primaryActionIds={primaryActionIds}
        hasSettings={Boolean(catalogClient)}
        onOpenSettings={() => setSettingsOpen(true)}
        onToggleChip={onToggle}
        onRevealChip={(id) => void controller.reveal(id)}
        onRefreshContext={() => void controller.refreshContext()}
        onInvokeSkill={(name, args) => void controller.invokeSkill(name, args)}
        onQuickAction={onQuickAction}
      />

      <dialog
        ref={settingsDialogRef}
        className="ge-modal"
        aria-label="Catalog and routing settings"
        onClose={() => setSettingsOpen(false)}
        onCancel={() => setSettingsOpen(false)}
      >
        <div className="ge-modal-head">
          <span className="ge-modal-title">Catalog &amp; routing</span>
          <button
            type="button"
            className="ge-modal-close"
            aria-label="Close settings"
            onClick={() => setSettingsOpen(false)}
          >
            <span aria-hidden="true">✕</span>
          </button>
        </div>
        <GeminiCatalogPanel
          catalogClient={catalogClient}
          open={settingsOpen}
          disabled={actionBlocked}
          onApply={(selection: GeminiCatalogSelection) => {
            onCatalogRouting?.(applyCatalogSelection(selection));
          }}
        />
      </dialog>

      {state.suggestions.length > 0 && (
        <section className="suggestions" aria-label="Suggestions">
          {state.suggestions.map((s) => (
            <button
              key={s.id}
              type="button"
              className="suggestion"
              disabled={actionBlocked}
              onClick={() => {
                if (s.query) controller.onAutomate(s.query);
                controller.dismissSuggestion(s.id);
              }}
            >
              <span className="s-title">{s.title}</span>
              {s.detail && <span className="s-detail">{s.detail}</span>}
            </button>
          ))}
        </section>
      )}

      <main className="thread-region" aria-label="Conversation and activity">
        <MessageThread messages={state.messages} surface={surface} />

        <RunSteps steps={state.steps} />

        <ProposalCard
          proposals={state.proposals}
          onApply={(id: ChangeId) => void controller.applyProposal(id)}
        />

        {state.error && (
          <div className="panel-error" role="alert">
            {state.error}
          </div>
        )}
      </main>

      {hasBlockingGate && (
        <section className="gate-rail" aria-label="Decision required">
          <CommandPlanCard
            pending={state.pendingCommandPlan}
            onConfirm={() => controller.confirmCommandPlan()}
            onCancel={() => controller.cancelCommandPlan()}
          />

          <PlanApprovalCard
            plan={state.pendingPlan}
            onApprove={() => controller.approvePlan()}
            onReject={() => controller.rejectPlan()}
          />

          <WriteApprovalCard
            pending={state.pendingWrite}
            onApprove={() => controller.approvePendingWrite()}
            onReject={() => controller.rejectPendingWrite()}
          />
        </section>
      )}

      <QuickActionParamForm
        key={paramFill?.id}
        action={paramFill}
        onSubmit={onParamFillSubmit}
        onCancel={() => setParamFill(undefined)}
      />

      <Composer
        busy={state.busy}
        disabled={hasBlockingGate}
        surface={surface}
        allowedIntents={allowedIntents}
        // The plain-string fallback (unreachable while `onInvoke` is set, since Composer prefers it).
        // Still guarded against an unfilled `{{…}}` so this is not a second un-guarded `send` seam if a
        // refactor ever drops `onInvoke` (security review H, LOW).
        onSend={(q) => {
          if (!hasUnfilledPlaceholder(q)) void controller.send(q);
        }}
        onCancel={() => controller.cancel()}
        onInvoke={onInvoke}
        placeholder={SURFACE_PLACEHOLDER[surface]}
      />
    </div>
  );
}
