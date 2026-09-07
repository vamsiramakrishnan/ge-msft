import {
  RuntimeHooks,
  completedEffectsExtension,
  installRuntimeExtensions,
  type RuntimeExtension,
} from '@ge/runtime';
import { TriggerRegistry } from '@ge/triggers';

/** Add trusted, bundled tenant extensions here. Never load executable hooks from documents or remote config. */
export const APPLICATION_EXTENSIONS: readonly RuntimeExtension[] = [
  completedEffectsExtension,
  {
    id: 'core.context',
    setup(api) {
      api.trigger({
        id: 'active-message',
        on: 'mail-received',
        handle: () => ({
          kind: 'suggest',
          title: 'Catch up on this message',
          detail: 'Summarize the request and identify open commitments.',
          query:
            '/summarize @this Summarize this message and identify explicit commitments and unanswered questions.',
        }),
      });
      api.trigger({
        id: 'meeting-ended',
        on: 'meeting-ended',
        handle: () => ({
          kind: 'suggest',
          title: 'Review decisions and owners',
          detail: 'Separate agreed actions from proposals.',
          query:
            '/ask @this Identify recorded decisions, owners, due dates, and unresolved questions. Distinguish proposals from agreements.',
        }),
      });
    },
  },
];

export function createApplicationRuntime(
  extensions: readonly RuntimeExtension[] = APPLICATION_EXTENSIONS,
) {
  const hooks = new RuntimeHooks();
  const triggers = new TriggerRegistry();
  const dispose = installRuntimeExtensions(extensions, { hooks, triggers });
  return { hooks, triggers, dispose };
}
