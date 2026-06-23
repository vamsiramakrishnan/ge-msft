import { useEffect } from 'react';
import type { Surface, ChangeId } from '@ge/contracts';
import type { PanelController } from '../../controller.js';
import { usePanelState } from '../usePanelState.js';
import { ContextTray } from './ContextTray.js';
import { MessageThread } from './MessageThread.js';
import { Composer } from './Composer.js';
import { ProposalCard } from './ProposalCard.js';
import { RunSteps } from './RunSteps.js';
import { WriteApprovalCard } from './WriteApprovalCard.js';
import { PlanApprovalCard } from './PlanApprovalCard.js';
import { SkillsPanel } from './SkillsPanel.js';

export interface AppProps {
  controller: PanelController;
  surface: Surface;
  agentLabel?: string;
}

const SURFACE_PLACEHOLDER: Readonly<Record<string, string>> = {
  word: 'Ask about the selection…',
  excel: 'Ask about this range…',
  outlook: 'Ask about this email…',
  teams: 'Ask about this meeting…',
};

/**
 * The task pane. A thin React view over `PanelController` state: header (agent identity), context
 * tray (attach/detach chips), streamed grounded thread with citations, proposal-review cards, and
 * the composer (send / cancel). No host or network code here — the controller owns all of that.
 */
export function App({ controller, surface, agentLabel }: AppProps): JSX.Element {
  const state = usePanelState(controller);

  // Load the attachable-context chips once on mount.
  useEffect(() => {
    void controller.refreshContext();
  }, [controller]);

  const onToggle = (id: string, attach: boolean): void => {
    if (attach) void controller.attach(id);
    else controller.detach(id);
  };

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

      <Composer
        busy={state.busy}
        onSend={(q) => void controller.send(q)}
        onRun={(t) => void controller.runCommands(t)}
        onCancel={() => controller.cancel()}
        placeholder={SURFACE_PLACEHOLDER[surface]}
      />
    </div>
  );
}
