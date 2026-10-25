export interface NativeReadOptions {
  signal?: AbortSignal;
  /** Absolute operation deadline shared by a sequence of read-only captures. */
  deadline?: number;
}

/** Read-only native callbacks may stall. Late completion cannot regain execution authority. */
export async function boundedNativeRead<T>(
  read: () => Promise<T>,
  options: NativeReadOptions = {},
): Promise<T> {
  if (options.signal?.aborted) throw new Error('Native inspection was cancelled.');
  const remaining = Math.min(2000, (options.deadline ?? Date.now() + 2000) - Date.now());
  if (remaining <= 0) throw new Error('Native inspection exceeded its read deadline.');
  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const finish = (value: T | undefined, error?: unknown): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      options.signal?.removeEventListener('abort', abort);
      if (error !== undefined) reject(error);
      else resolve(value as T);
    };
    const abort = () => finish(undefined, new Error('Native inspection was cancelled.'));
    const timer = setTimeout(
      () => finish(undefined, new Error('Native inspection exceeded its read deadline.')),
      remaining,
    );
    options.signal?.addEventListener('abort', abort, { once: true });
    if (options.signal?.aborted) {
      abort();
      return;
    }
    Promise.resolve()
      .then(() => {
        if (settled) throw new Error('Native inspection was cancelled before capture.');
        return read();
      })
      .then(
        (value) => finish(value),
        (error) => finish(undefined, error ?? new Error('Native inspection failed.')),
      );
  });
}
