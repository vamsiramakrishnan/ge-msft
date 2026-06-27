/**
 * Trailing-edge debounce with injectable timers (so it's deterministically testable).
 * Event sources wrap high-frequency emissions (selection moves, keystrokes) with this so
 * triggers fire once the user pauses, not on every micro-event.
 */
export interface Scheduler {
  set(fn: () => void, ms: number): unknown;
  clear(handle: unknown): void;
}

const defaultScheduler: Scheduler = {
  set: (fn, ms) => setTimeout(fn, ms),
  clear: (h) => clearTimeout(h as ReturnType<typeof setTimeout>),
};

export interface Debounced<A extends unknown[]> {
  (...args: A): void;
  cancel(): void;
  flush(): void;
}

export function debounce<A extends unknown[]>(
  fn: (...args: A) => void,
  ms: number,
  scheduler: Scheduler = defaultScheduler,
): Debounced<A> {
  let handle: unknown = null;
  let pending: A | null = null;

  const run = (): void => {
    handle = null;
    if (pending) {
      const args = pending;
      pending = null;
      fn(...args);
    }
  };

  const debounced = ((...args: A): void => {
    pending = args;
    if (handle !== null) scheduler.clear(handle);
    handle = scheduler.set(run, ms);
  }) as Debounced<A>;

  debounced.cancel = (): void => {
    if (handle !== null) scheduler.clear(handle);
    handle = null;
    pending = null;
  };
  debounced.flush = (): void => {
    if (handle !== null) {
      scheduler.clear(handle);
      run();
    }
  };
  return debounced;
}
