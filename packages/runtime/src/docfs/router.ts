// packages/runtime/src/docfs/router.ts
import {
  parseDocPath,
  type DirEntry,
  type DocFs,
  type FileStat,
  type FileView,
  type ReadOpts,
  type SearchMatch,
  type SearchOpts,
} from '@ge/contracts';
import type { Mount } from './mount.js';

/** Dispatches DocFs calls to the Mount that owns the path's first segment. */
export class DocFsRouter implements DocFs {
  private readonly mounts = new Map<string, Mount>();
  constructor(mounts: Mount[]) {
    for (const m of mounts) this.mounts.set(m.prefix, m);
  }

  private resolve(path: string): { mount: Mount; rel: string; abs: (rel: string) => string } {
    const { mount, rel } = parseDocPath(path);
    const m = this.mounts.get(mount);
    if (!m) throw new Error(`DocFs: unknown mount '${mount}' in ${path}`);
    const abs = (r: string) => `/${mount}${r ? `/${r}` : ''}`;
    return { mount: m, rel, abs };
  }

  async readdir(path: string): Promise<DirEntry[]> {
    if (path === '/' || path === '') {
      return [...this.mounts.keys()].sort().map((name) => ({ name, kind: 'dir' as const }));
    }
    const { mount, rel } = this.resolve(path);
    return mount.readdir(rel);
  }

  async stat(path: string): Promise<FileStat | null> {
    if (path === '/' || path === '') return { path: '/', kind: 'dir', size: 0 };
    const { mount, rel, abs } = this.resolve(path);
    const s = await mount.stat(rel);
    return s ? { ...s, path: abs(rel) } : null;
  }

  async readFile(path: string, opts?: ReadOpts): Promise<FileView | null> {
    const { mount, rel, abs } = this.resolve(path);
    const v = await mount.readFile(rel, opts);
    return v ? { ...v, path: abs(rel) } : null;
  }

  async search(path: string, pattern: string, opts?: SearchOpts): Promise<SearchMatch[]> {
    const { mount, rel, abs } = this.resolve(path);
    const matches = await mount.search(rel, pattern, opts);
    // Mounts return match paths relative to the mount; stamp them absolute.
    return matches.map((m) => ({ ...m, path: abs(m.path) }));
  }
}
