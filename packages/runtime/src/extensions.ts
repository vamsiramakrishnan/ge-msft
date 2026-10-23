import { assessActuationResult, type ActuationResult } from '@ge/contracts';
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

/**
 * An actual verifier: a parsed `done` does not make failed effects or an exhausted loop successful.
 * A policy-'blocked' task can still carry failed (not uncertain) effects, so a guard that cares
 * about writes should read `outcome.effects`, not the status alone.
 */
export const completedEffectsExtension: RuntimeExtension = {
  id: 'core.outcomes',
  setup(api) {
    api.on({
      id: 'verify',
      on: 'task:verify',
      mode: 'guard',
      handle({ outcome }) {
        if (outcome.status === 'incomplete') {
          // Only warn that changes may have applied when a write ran or its outcome is unknown; a
          // task stopped before any write landed says so (live 2026-10-14: "0 changes applied").
          const mayHaveApplied = outcome.effects.some((result) =>
            // A read-only view of the receipt; the assessment only reads it.
            ['verified', 'unverified', 'uncertain'].includes(
              assessActuationResult(result as ActuationResult),
            ),
          );
          return {
            kind: 'block',
            reason: mayHaveApplied
              ? 'The task did not complete all its operations. Review the execution results before continuing. Some changes may already have applied.'
              : 'The task did not complete all its operations. No changes were applied.',
          };
        }
        return { kind: 'continue' };
      },
    });
  },
};
