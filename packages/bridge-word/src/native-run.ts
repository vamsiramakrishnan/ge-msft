import { observeNativeBatch } from '@ge/runtime';

export function nativeRun<T>(work: (context: Word.RequestContext) => Promise<T>): Promise<T> {
  return Word.run((context) => observeNativeBatch('word', context, work));
}
