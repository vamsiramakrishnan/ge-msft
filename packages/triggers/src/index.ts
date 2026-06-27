/**
 * @ge/triggers — the event-driven layer. Lifecycle events (host activity, estate changes,
 * and the pre/post-actuation hooks) flow through an EventBus to a TriggerRegistry that
 * matches them to triggers and returns outcomes (continue/block/suggest/automate). This is
 * how the add-in *reacts* and *triggers on* events, the way Claude Code uses hooks.
 */
export * from './event.js';
export { EventBus, type EventListener, type Unsubscribe } from './bus.js';
export { TriggerRegistry, type Trigger } from './registry.js';
export { debounce, type Debounced, type Scheduler } from './debounce.js';
