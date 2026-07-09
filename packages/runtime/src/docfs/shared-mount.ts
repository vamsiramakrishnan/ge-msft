// packages/runtime/src/docfs/shared-mount.ts
import type {
  DirEntry,
  FileStat,
  FileView,
  ReadOpts,
  SearchMatch,
  SearchOpts,
} from '@ge/contracts';
import { byteLength, truncateToBytes } from './bytes.js';
import type { Mount } from './mount.js';

const DEFAULT_MAX_BYTES = 256 * 1024;

/**
 * The narrow port `/shared` is built against — list/read/write/remove over named text artifacts.
 * `packages/runtime` never imports Graph directly: the concrete Graph-backed implementation
 * (`GraphSharedStore`, backed by `Files.ReadWrite.AppFolder`) lives in `@ge/graph-client` and is
 * wired in by the caller, exactly like `DocBridge`/`AuthClient` are injected rather than owned.
 */
export interface SharedStore {
  list(): Promise<{ name: string; size: number }[]>;
  read(path: string): Promise<string | undefined>;
  write(path: string, content: string): Promise<void>;
  remove(path: string): Promise<void>;
}

/**
 * `/shared` — the cross-surface handoff store. A value `share`d from one surface's session (e.g.
 * Excel) is readable back by name from any other surface's session (Word, PowerPoint, Teams)
 * through this same mount, the same way `/work` makes `save`d artifacts `cat`/`grep`-able. `/` in
 * a name is a pseudo-directory (naming convention only), grouped the same way `/work` and `/skills`
 * group `/`-containing names — no new persisted state.
 */
export function sharedMount(store: SharedStore): Mount {
  return {
    prefix: 'shared',
    async readdir(rel: string): Promise<DirEntry[]> {
      const prefix = rel === '' ? '' : `${rel}/`;
      const dirs = new Set<string>();
      const entries: DirEntry[] = [];
      for (const item of await store.list()) {
        if (!item.name.startsWith(prefix)) continue;
        const remainder = item.name.slice(prefix.length);
        const slash = remainder.indexOf('/');
        if (slash === -1) {
          entries.push({ name: remainder, kind: 'file', size: item.size });
        } else {
          dirs.add(remainder.slice(0, slash));
        }
      }
      for (const name of dirs) entries.push({ name, kind: 'dir' });
      return entries;
    },
    async stat(rel: string): Promise<FileStat | null> {
      const items = await store.list();
      const item = items.find((i) => i.name === rel);
      if (item) return { path: '', kind: 'file', size: item.size };
      const hasChildren = items.some((i) => i.name.startsWith(`${rel}/`));
      return hasChildren ? { path: '', kind: 'dir', size: 0 } : null;
    },
    async readFile(rel: string, opts?: ReadOpts): Promise<FileView | null> {
      const content = await store.read(rel);
      if (content === undefined) return null;
      const max = opts?.maxBytes ?? DEFAULT_MAX_BYTES;
      const { text, truncated } = truncateToBytes(content, max);
      return { path: '', text, bytes: byteLength(text), truncated };
    },
    async search(rel: string, pattern: string, opts?: SearchOpts): Promise<SearchMatch[]> {
      if (!pattern) return [];
      const max = opts?.max ?? 50;
      const needle = pattern.toLowerCase();
      const prefix = rel === '' ? '' : `${rel}/`;
      const matches: SearchMatch[] = [];
      for (const item of await store.list()) {
        if (!item.name.startsWith(prefix)) continue;
        const content = await store.read(item.name);
        if (content === undefined) continue;
        const lines = content.split('\n');
        for (let i = 0; i < lines.length; i++) {
          if (lines[i]!.toLowerCase().includes(needle)) {
            matches.push({ path: item.name, line: i + 1, text: lines[i]! });
            if (matches.length >= max) return matches;
          }
        }
      }
      return matches;
    },
  };
}
