import { observeNativeBatch } from '@ge/runtime';
import { isSet } from './capabilities-runtime.js';

/**
 * Options for one Office.js batch.
 *
 * `undoGroup` asks the host to merge every request in the batch into ONE undo entry
 * (`Excel.RunOptions.mergeUndoGroup`, ExcelApi 1.20) so an approved plan is a single Ctrl+Z for
 * the user. On a host below 1.20 the option is simply not passed — the batch still runs; the
 * host records its own per-request undo entries as before. Read-only batches never ask for it.
 */
export interface NativeRunOptions {
  undoGroup?: boolean;
}

/** True iff the host can merge a batch into one undo group (ExcelApi 1.20). */
export function supportsUndoGroup(): boolean {
  return isSet('ExcelApi', '1.20');
}

type Batch<T> = (context: Excel.RequestContext) => Promise<T>;

export function nativeRun<T>(work: Batch<T>): Promise<T>;
export function nativeRun<T>(options: NativeRunOptions, work: Batch<T>): Promise<T>;
export function nativeRun<T>(
  optionsOrWork: NativeRunOptions | Batch<T>,
  maybeWork?: Batch<T>,
): Promise<T> {
  const options = typeof optionsOrWork === 'function' ? {} : optionsOrWork;
  const work = typeof optionsOrWork === 'function' ? optionsOrWork : maybeWork!;
  const batch: Batch<T> = (context) => observeNativeBatch('excel', context, work);
  if (options.undoGroup && supportsUndoGroup()) {
    return Excel.run({ mergeUndoGroup: true }, batch);
  }
  return Excel.run(batch);
}
