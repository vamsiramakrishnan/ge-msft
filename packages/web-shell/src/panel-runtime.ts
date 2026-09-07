import { Orchestrator, type AssistSession, type DocBridge } from '@ge/runtime';
import type { HostEvent, TriggerRegistry } from '@ge/triggers';
import type { PanelController } from './controller.js';

/** Own the event/hook subscriptions independently of React renders and authentication retries. */
export function connectPanelRuntime(input: {
  session: AssistSession;
  bridge: DocBridge;
  controller: PanelController;
  triggers: TriggerRegistry;
}) {
  const { session, bridge, controller, triggers } = input;
  let disposed = false;
  let lifetime = new AbortController();
  let refreshPending = false;
  let refreshing = false;
  const flushContext = async (): Promise<void> => {
    const state = controller.getState();
    if (
      disposed ||
      lifetime.signal.aborted ||
      refreshing ||
      !refreshPending ||
      state.busy ||
      state.pendingPlan ||
      state.pendingWrite ||
      state.pendingShare ||
      state.pendingCommandPlan
    )
      return;
    refreshing = true;
    refreshPending = false;
    try {
      await controller.refreshContext({ whenIdle: true, signal: lifetime.signal });
    } finally {
      refreshing = false;
      const current = controller.getState();
      if (
        current.busy ||
        current.pendingPlan ||
        current.pendingWrite ||
        current.pendingShare ||
        current.pendingCommandPlan
      )
        refreshPending = true;
      if (refreshPending && !disposed) void flushContext();
    }
  };
  const orchestrator = new Orchestrator(
    bridge,
    triggers,
    {
      async onContext(event, { signal }) {
        await session.hooks.run(
          'host:event',
          { event },
          { taskId: 'host-events', surface: bridge.surface, signal },
        );
        if (disposed || signal.aborted) return;
        await controller.onContext(event);
        if (
          [
            'selection-changed',
            'document-changed',
            'mail-received',
            'mail-compose',
            'estate-changed',
          ].includes(event.type)
        ) {
          refreshPending = true;
          void flushContext();
        }
      },
      onSuggest: controller.onSuggest,
      // Automatic trigger outcomes become explicit invitations in the production app. Accepting
      // one still traverses the message hooks, normal routing, and any subsequent write approval.
      onAutomate(query) {
        controller.onSuggest({ title: 'Suggested next step', query });
      },
      onError(code) {
        controller.onRuntimeNotice(`Event processing: ${code}`);
      },
    },
    { emitLifecycle: true },
  );
  const offState = controller.subscribe(() => {
    if (refreshPending) void flushContext();
  });
  const offHooks = session.hooks.subscribe((record) => {
    if (['error', 'timeout'].includes(record.outcome))
      controller.onRuntimeNotice(
        `Extension ${record.hookId}: ${record.outcome} at ${record.phase}.`,
      );
  });
  const offTriggers = triggers.subscribe((record) =>
    controller.onRuntimeNotice(`Extension ${record.triggerId}: check failed at ${record.event}.`),
  );
  const offRuns = session.executions.subscribe((run) => {
    if (run.status !== 'running')
      controller.onRuntimeNotice(
        `Task ${run.status}. ${run.effects.filter((e) => e.ok).length} changes applied; ${run.effects.filter((e) => !e.ok).length} not applied.`,
      );
  });
  const stop = (): void => {
    lifetime.abort();
    refreshPending = false;
    orchestrator.stop();
  };
  return {
    start: () => {
      if (!disposed) {
        if (lifetime.signal.aborted) lifetime = new AbortController();
        orchestrator.start();
      }
    },
    stop,
    receive: (event: HostEvent) => orchestrator.publish(event),
    idle: () => orchestrator.idle(),
    dispose() {
      if (disposed) return;
      disposed = true;
      stop();
      offState();
      offHooks();
      offRuns();
      offTriggers();
    },
  };
}
