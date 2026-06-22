import { useSyncExternalStore } from 'react';
import type { PanelController, PanelState } from '../controller.js';

/**
 * Subscribe a React tree to the `PanelController`'s immutable `PanelState`. The controller already
 * exposes `subscribe(listener)` returning an unsubscribe and `getState()` returning the current
 * snapshot, which is exactly the `useSyncExternalStore` contract — so React re-renders on every
 * `set()` without the panel logic knowing React exists.
 */
export function usePanelState(controller: PanelController): PanelState {
  return useSyncExternalStore(
    (onChange) => controller.subscribe(onChange),
    () => controller.getState(),
    () => controller.getState(),
  );
}
