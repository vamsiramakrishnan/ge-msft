import { useEffect } from 'react';
import {
  deriveOutput,
  GROUND_SOURCE_TO_SELECTION_KIND,
  type Surface,
  type ChangeId,
  type QuickAction,
  type Intent,
  type GroundingSelection,
} from '@ge/contracts';
import {
  resolveGrounding,
  type GroundingResolveContext,
  type ResolvedGrounding,
} from '@ge/gemini-client';
import type { PanelController } from '../../controller.js';
import { usePanelState } from '../usePanelState.js';
import { ContextTray } from './ContextTray.js';
import { MessageThread } from './MessageThread.js';
import { Composer, type ComposerInvocation, type ComposerMention } from './Composer.js';
import { QuickActionBar } from './QuickActionBar.js';
import { invocationToSeed, quickActionToInvocation } from './quick-action-seed.js';
import { ProposalCard } from './ProposalCard.js';
import { RunSteps } from './RunSteps.js';
import { WriteApprovalCard } from './WriteApprovalCard.js';
import { PlanApprovalCard } from './PlanApprovalCard.js';
import { CommandPlanCard } from './CommandPlanCard.js';
import { SkillsPanel } from './SkillsPanel.js';

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
}

const SURFACE_PLACEHOLDER: Readonly<Record<string, string>> = {
  word: 'Ask about the selection…',
  excel: 'Ask about this range…',
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
export function App({ controller, surface, agentLabel, allowedIntents }: AppProps): JSX.Element {
  const state = usePanelState(controller);

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
    const seed = invocationToSeed(inv);
    const grounding = invocationToGrounding(inv);
    // EXPERIENCE.md §F — the planner-confirm front door, enforced ONLY for complex free-text: a
    // composer-typed (raw ≠ '') actuating instruction with constraints first proposes a confirmable
    // CommandPlan. A chip/preset (raw === '') or a simple instruction routes straight to the executor.
    const composerOrigin = inv.raw.trim() !== '';
    if (composerOrigin && isActuating(inv.intent) && isComplexInstruction(inv.instruction)) {
      void controller.proposePlan(seed, grounding);
    } else if (isActuating(inv.intent)) {
      void controller.runCommands(seed, grounding);
    } else {
      void controller.send(seed, grounding);
    }
  };

  // A chip is the same typed Invocation a composer line is — pre-filled, then dispatched through the
  // one shared `dispatch`. (`action.output` and `isActuating` agree by construction, ADR-0006.)
  const onQuickAction = (action: QuickAction): void => dispatch(quickActionToInvocation(action));

  // A structured composer submit (`/verb` intent + scope + `@`-mentions + instruction) → the same path.
  const onInvoke = (inv: ComposerInvocation): void => dispatch(inv);

  return (
    <div className="panel" data-surface={surface} aria-busy={state.busy}>
      <header className="ph">
        <div className="pht">
          <div className="av" aria-hidden="true" />
          <div>
            <div className="pn">Gemini Enterprise</div>
            <div className="pss">{agentLabel ?? 'Grounded on your research unit'}</div>
          </div>
        </div>
      </header>

      <ContextTray
        chips={state.chips}
        onToggle={onToggle}
        onRefresh={() => void controller.refreshContext()}
      />

      {state.suggestions.length > 0 && (
        <section className="suggestions" aria-label="Suggestions">
          {state.suggestions.map((s) => (
            <button
              key={s.id}
              type="button"
              className="suggestion"
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

      <SkillsPanel
        skills={state.skills ?? []}
        onInvoke={(name, args) => void controller.invokeSkill(name, args)}
      />

      <main className="thread-region" aria-label="Conversation and activity">
        <MessageThread messages={state.messages} />

        <RunSteps steps={state.steps} />

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

      <QuickActionBar
        surface={surface}
        allowedIntents={allowedIntents}
        busy={state.busy}
        onAction={onQuickAction}
      />

      <Composer
        busy={state.busy}
        surface={surface}
        allowedIntents={allowedIntents}
        onSend={(q) => void controller.send(q)}
        onCancel={() => controller.cancel()}
        onInvoke={onInvoke}
        placeholder={SURFACE_PLACEHOLDER[surface]}
      />
    </div>
  );
}
