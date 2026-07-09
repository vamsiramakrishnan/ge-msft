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
    async readdir(): Promise<DirEntry[]> {
      return store.list().map((a) => ({ name: a.name, kind: 'file' as const, size: a.bytes }));
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
