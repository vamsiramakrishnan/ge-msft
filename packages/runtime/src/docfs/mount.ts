// packages/runtime/src/docfs/mount.ts
import type {
  DirEntry,
  FileStat,
  FileView,
  ReadOpts,
  SearchMatch,
  SearchOpts,
} from '@ge/contracts';

/** A DocFs mount: the four read/search ops over paths RELATIVE to the mount (no leading mount name). */
export interface Mount {
  readonly prefix: string;
  readonly readonly?: boolean;
  readdir(rel: string): Promise<DirEntry[]>;
  stat(rel: string): Promise<FileStat | null>;
  readFile(rel: string, opts?: ReadOpts): Promise<FileView | null>;
  search(rel: string, pattern: string, opts?: SearchOpts): Promise<SearchMatch[]>;
}

/** Mark a mount read-only. Reads pass through unchanged; the flag documents intent for callers/UI. */
export function readOnly(mount: Mount): Mount {
  return { ...mount, readonly: true };
}
