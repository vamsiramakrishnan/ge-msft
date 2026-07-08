// packages/contracts/src/docfs.ts
/**
 * DocFs — the conductor's read/search narrow waist. The live document, the workspace, and (later)
 * sources/skills/peers are all presented as one POSIX-style filesystem the model navigates. Writes
 * are NOT part of this interface — mutation stays on the gated `DocBridge.actuate` path.
 */
export interface DirEntry {
  name: string;
  kind: 'file' | 'dir';
  size?: number;
}
export interface FileStat {
  path: string;
  kind: 'file' | 'dir';
  size: number;
}
export interface FileView {
  path: string;
  text: string;
  bytes: number;
  truncated: boolean;
}
export interface SearchMatch {
  path: string;
  line: number;
  text: string;
}
export interface ReadOpts {
  maxBytes?: number;
}
export interface SearchOpts {
  max?: number;
  context?: number;
}

export interface DocFs {
  readdir(path: string): Promise<DirEntry[]>;
  stat(path: string): Promise<FileStat | null>;
  readFile(path: string, opts?: ReadOpts): Promise<FileView | null>;
  search(path: string, pattern: string, opts?: SearchOpts): Promise<SearchMatch[]>;
}

/** Split an absolute `/<mount>/<rest>` path; normalize `.`/`..` within the mount; reject escapes. */
export function parseDocPath(path: string): { mount: string; rel: string } {
  if (!path.startsWith('/')) throw new Error(`DocFs path must be absolute: ${path}`);
  const segments = path.split('/').filter((s) => s.length > 0);
  if (segments.length === 0) throw new Error('DocFs path must include a mount, e.g. /doc');
  const [mount, ...rest] = segments;
  const stack: string[] = [];
  for (const s of rest) {
    if (s === '.') continue;
    if (s === '..') {
      if (stack.length === 0) throw new Error(`DocFs path escapes its mount: ${path}`);
      stack.pop();
      continue;
    }
    stack.push(s);
  }
  return { mount: mount!, rel: stack.join('/') };
}
