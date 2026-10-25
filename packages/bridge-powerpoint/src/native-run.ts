import { observeNativeBatch } from '@ge/runtime';

export function nativeRun<T>(work: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T> {
  return PowerPoint.run((context) => observeNativeBatch('powerpoint', context, work));
}
