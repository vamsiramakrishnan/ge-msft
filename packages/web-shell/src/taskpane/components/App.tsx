import { useEffect } from 'react';
import {
  deriveOutput,
  type Surface,
  type ChangeId,
  type QuickAction,
  type Intent,
} from '@ge/contracts';
import type { PanelController } from '../../controller.js';
import { usePanelState } from '../usePanelState.js';
import { ContextTray } from './ContextTray.js';
import { MessageThread } from './MessageThread.js';
import { Composer, type ComposerInvocation } from './Composer.js';
import { QuickActionBar } from './QuickActionBar.js';
import { invocationToSeed, quickActionToInvocation } from './quick-action-seed.js';
import { ProposalCard } from './ProposalCard.js';
import { RunSteps } from './RunSteps.js';
import { WriteApprovalCard } from './WriteApprovalCard.js';
import { PlanApprovalCard } from './PlanApprovalCard.js';
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
  // The controller still takes a string task, so the typed tuple is rendered to its deterministic
  // seed only at this seam. No new gate is introduced — chips/`/verbs` only seed the existing route.
  const dispatch = (inv: ComposerInvocation): void => {
    const seed = invocationToSeed(inv);
    if (isActuating(inv.intent)) void controller.runCommands(seed);
    else void controller.send(seed);
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
