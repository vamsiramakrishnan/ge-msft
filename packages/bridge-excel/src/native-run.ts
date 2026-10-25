import { observeNativeBatch } from '@ge/runtime';

export function nativeRun<T>(work: (context: Excel.RequestContext) => Promise<T>): Promise<T> {
  return Excel.run((context) => observeNativeBatch('excel', context, work));
}
