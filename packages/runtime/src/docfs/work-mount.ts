// packages/runtime/src/docfs/work-mount.ts
import type { DirEntry, FileView, ReadOpts, SearchMatch, SearchOpts } from '@ge/contracts';
import type { WorkspaceStore } from '../workspace.js';
import { byteLength, truncateToBytes } from './bytes.js';
import type { Mount } from './mount.js';

const DEFAULT_MAX_BYTES = 256 * 1024;

/** `/work` — the in-memory WorkspaceStore artifacts as files. Reshape, not rewrite: this delegates
 *  to the store's existing list/get/grep so there is one artifact store, behind the DocFs waist. */
export function workMount(store: WorkspaceStore): Mount {
  return {
    prefix: 'work',
    async readdir(rel: string): Promise<DirEntry[]> {
      // Artifact names are flat strings in WorkspaceStore; a `/` in a name is rendered here as a
      // pseudo-directory — naming-convention only, no new persisted state. Partition by whether a
      // name sits directly under `rel` (a file) or has a further segment past it (collapse to one
      // `dir` entry per distinct next segment, deduped).
      const prefix = rel === '' ? '' : `${rel}/`;
      const dirs = new Set<string>();
      const entries: DirEntry[] = [];
      for (const a of store.list()) {
        if (!a.name.startsWith(prefix)) continue;
        const remainder = a.name.slice(prefix.length);
        const slash = remainder.indexOf('/');
        if (slash === -1) {
          entries.push({ name: remainder, kind: 'file' as const, size: a.bytes });
        } else {
          dirs.add(remainder.slice(0, slash));
        }
      }
      for (const name of dirs) entries.push({ name, kind: 'dir' as const });
      return entries;
    },
    async stat(rel) {
      const artifact = store.get(rel);
      return artifact ? { path: '', kind: 'file' as const, size: artifact.bytes } : null;
    },
    async readFile(rel, opts?: ReadOpts): Promise<FileView | null> {
      const artifact = store.get(rel);
      if (!artifact) return null;
      const max = opts?.maxBytes ?? DEFAULT_MAX_BYTES;
      const { text, truncated } = truncateToBytes(artifact.text, max);
      return {
        path: '',
        text,
        bytes: byteLength(text),
        truncated,
      };
    },
    async search(rel, pattern, opts?: SearchOpts): Promise<SearchMatch[]> {
      const max = opts?.max ?? 50;
      const context = opts?.context ?? 0;
      if (rel) {
        const res = store.grep(rel, pattern, context);
        if (res.workspace !== 'grep') return [];
        return res.matches.slice(0, max).map((m) => ({ path: rel, line: m.line, text: m.text }));
      }
      const hits: SearchMatch[] = [];
      for (const artifact of store.list()) {
        const res = store.grep(artifact.name, pattern, context);
        if (res.workspace !== 'grep') continue;
        for (const m of res.matches) {
          hits.push({ path: artifact.name, line: m.line, text: m.text });
        }
      }
      return hits.slice(0, max);
    },
  };
}
