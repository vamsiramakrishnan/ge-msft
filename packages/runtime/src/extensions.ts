import type { Trigger, TriggerRegistry } from '@ge/triggers';
import type { RuntimeHook, RuntimeHookPhase, RuntimeHooks } from './hooks.js';

export interface RuntimeExtensionApi {
  on<K extends RuntimeHookPhase>(hook: RuntimeHook<K>): void;
  trigger(trigger: Trigger): void;
}
export interface RuntimeExtension {
  id: string;
  /** Synchronous registration only. Async work belongs in deadline-bound handlers. */
  setup(api: RuntimeExtensionApi): void;
}

/** All-or-nothing installation with namespaced IDs; disposing removes only this installation. */
export function installRuntimeExtensions(
  extensions: readonly RuntimeExtension[],
  services: { hooks: RuntimeHooks; triggers: TriggerRegistry },
): () => void {
  const ids = new Set<string>();
  const disposers: Array<() => void> = [];
  let closed = false;
  const dispose = (): void => {
    closed = true;
    for (const off of disposers.splice(0).reverse()) off();
  };
  try {
    for (const extension of extensions) {
      if (!/^[a-zA-Z0-9][a-zA-Z0-9._-]{0,59}$/.test(extension.id) || ids.has(extension.id))
        throw new Error('Invalid or duplicate runtime extension id');
      ids.add(extension.id);
      let registering = true;
      const check = (): void => {
        if (closed || !registering) throw new Error('Extension registration is closed');
      };
      try {
        const result: unknown = extension.setup({
          on(hook) {
            check();
            disposers.push(services.hooks.register({ ...hook, id: `${extension.id}/${hook.id}` }));
          },
          trigger(trigger) {
            check();
            disposers.push(
              services.triggers.register({ ...trigger, id: `${extension.id}/${trigger.id}` }),
            );
          },
        });
        if (result && typeof result === 'object' && 'then' in result) {
          void Promise.resolve(result).catch(() => {});
          throw new Error('Runtime extension setup must be synchronous');
        }
      } finally {
        registering = false;
      }
    }
    return dispose;
  } catch (error) {
    dispose();
    throw error;
  }
}

/** An actual verifier: a parsed `done` does not make failed effects or an exhausted loop successful. */
export const completedEffectsExtension: RuntimeExtension = {
  id: 'core.outcomes',
  setup(api) {
    api.on({
      id: 'verify',
      on: 'task:verify',
      mode: 'guard',
      handle({ outcome }) {
        if (outcome.status === 'incomplete')
          return {
            kind: 'block',
            reason:
              'The task did not complete all its operations. Review the execution results before continuing. Some changes may already have applied.',
          };
        return { kind: 'continue' };
      },
    });
  },
};
