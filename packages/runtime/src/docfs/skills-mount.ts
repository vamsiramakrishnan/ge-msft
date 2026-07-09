// packages/runtime/src/docfs/skills-mount.ts
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

const DEFAULT_MAX_BYTES = 64 * 1024;

/**
 * `/skills` — the mounted Gemini Enterprise skill bundles' own reference files (SKILL.md,
 * references/*.md), as a read-only DocFs index. Unlike `/doc`, there is no live host to lazily
 * query: `files` is a plain `{relativePath: content}` map the CALLER supplies, built at bundle time
 * from the same checked-in `skill/` sources this repo's own skillsSpec already mounts server-side
 * per turn — this mount just lets the model browse/re-fetch one specific reference file mid-turn
 * instead of relying on everything being crammed into the skill's own always-on context, mirroring
 * how the conductor's VFS makes `SKILL.md` a `cat`-able file rather than an opaque system-prompt
 * blob. `/` in a path is a real directory here (skill bundles ARE nested — `<bundle>/references/…`),
 * grouped the same way `/work`'s namespacing groups `/`-containing artifact names.
 */
export function skillsMount(files: Readonly<Record<string, string>>): Mount {
  const paths = Object.keys(files);

  return {
    prefix: 'skills',
    async readdir(rel: string): Promise<DirEntry[]> {
      const prefix = rel === '' ? '' : `${rel}/`;
      const dirs = new Set<string>();
      const entries: DirEntry[] = [];
      for (const p of paths) {
        if (!p.startsWith(prefix)) continue;
        const remainder = p.slice(prefix.length);
        const slash = remainder.indexOf('/');
        if (slash === -1) {
          entries.push({ name: remainder, kind: 'file', size: byteLength(files[p]!) });
        } else {
          dirs.add(remainder.slice(0, slash));
        }
      }
      for (const name of dirs) entries.push({ name, kind: 'dir' });
      return entries;
    },
    async stat(rel: string): Promise<FileStat | null> {
      const content = files[rel];
      if (content !== undefined) return { path: '', kind: 'file', size: byteLength(content) };
      const hasChildren = paths.some((p) => p.startsWith(`${rel}/`));
      return hasChildren ? { path: '', kind: 'dir', size: 0 } : null;
    },
    async readFile(rel: string, opts?: ReadOpts): Promise<FileView | null> {
      const content = files[rel];
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
      for (const p of paths) {
        if (!p.startsWith(prefix)) continue;
        const lines = files[p]!.split('\n');
        for (let i = 0; i < lines.length; i++) {
          if (lines[i]!.toLowerCase().includes(needle)) {
            matches.push({ path: p, line: i + 1, text: lines[i]! });
            if (matches.length >= max) return matches;
          }
        }
      }
      return matches;
    },
  };
}
