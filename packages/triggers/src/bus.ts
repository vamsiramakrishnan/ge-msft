import type { EventType, HostEvent } from './event.js';

export type EventListener = (event: HostEvent) => void;
export type Unsubscribe = () => void;

/**
 * A tiny synchronous event bus. Bridges/Graph `emit` host events; the orchestrator and the
 * trigger registry subscribe. Surface-agnostic and dependency-free.
 */
export class EventBus {
  private readonly listeners = new Map<EventType | '*', Set<EventListener>>();

  /** Subscribe to one event type, or '*' for all. Returns an unsubscribe fn. */
  on(type: EventType | '*', listener: EventListener): Unsubscribe {
    const set = this.listeners.get(type) ?? new Set<EventListener>();
    set.add(listener);
    this.listeners.set(type, set);
    return () => set.delete(listener);
  }

  emit(event: HostEvent): void {
    for (const l of this.listeners.get(event.type) ?? []) l(event);
    for (const l of this.listeners.get('*') ?? []) l(event);
  }

  clear(): void {
    this.listeners.clear();
  }
}
