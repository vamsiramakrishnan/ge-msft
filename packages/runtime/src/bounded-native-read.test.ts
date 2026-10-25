import { describe, expect, it, vi } from 'vitest';
import { boundedNativeRead } from './bounded-native-read.js';

describe('bounded native read-only callbacks', () => {
  it('expires stalled reads and ignores their later result', async () => {
    vi.useFakeTimers();
    try {
      let resolve!: (value: string) => void;
      const pending = boundedNativeRead(
        () =>
          new Promise<string>((finish) => {
            resolve = finish;
          }),
      );
      const check = expect(pending).rejects.toThrow('read deadline');
      await vi.advanceTimersByTimeAsync(2001);
      await check;
      resolve('too late');
      await Promise.resolve();
    } finally {
      vi.useRealTimers();
    }
  });
  it('cancels a read without granting its late native completion any authority', async () => {
    const controller = new AbortController();
    let finish!: (value: string) => void;
    const pending = boundedNativeRead(
      () =>
        new Promise<string>((resolve) => {
          finish = resolve;
        }),
      { signal: controller.signal },
    );
    await Promise.resolve();
    const check = expect(pending).rejects.toThrow('cancelled');
    controller.abort();
    await check;
    finish('late');
    await Promise.resolve();
  });
  it('honors an exhausted shared deadline without starting another capture', async () => {
    const read = vi.fn(async () => 'native');
    await expect(boundedNativeRead(read, { deadline: Date.now() - 1 })).rejects.toThrow('deadline');
    expect(read).not.toHaveBeenCalled();
  });
});
