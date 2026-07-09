import type { DocFs, SearchMatch } from '@ge/contracts';

export async function ls(fs: DocFs, path: string): Promise<string[]> {
  return (await fs.readdir(path)).map((e) => (e.kind === 'dir' ? `${e.name}/` : e.name));
}

export async function cat(
  fs: DocFs,
  path: string,
  opts?: { maxBytes?: number },
): Promise<{ text: string; truncated: boolean }> {
  const v = await fs.readFile(path, opts);
  if (!v) throw new Error(`cat: no such file: ${path}`);
  return { text: v.text, truncated: v.truncated };
}

export async function head(fs: DocFs, path: string, n = 10): Promise<{ lines: string[] }> {
  const v = await fs.readFile(path);
  if (!v) throw new Error(`head: no such file: ${path}`);
  return { lines: v.text.split('\n').slice(0, n) };
}

export async function grep(
  fs: DocFs,
  path: string,
  pattern: string,
  opts?: { max?: number; context?: number },
): Promise<{ matches: SearchMatch[] }> {
  return { matches: await fs.search(path, pattern, opts) };
}

export async function wc(fs: DocFs, path: string): Promise<{ lines: number; bytes: number }> {
  const v = await fs.readFile(path);
  if (!v) throw new Error(`wc: no such file: ${path}`);
  return {
    lines: v.text.split('\n').filter((l, i, a) => i < a.length - 1 || l.length > 0).length,
    bytes: v.bytes,
  };
}

/**
 * Recursive listing with an optional `*`/`?` glob on the leaf name. Every regex-special character
 * in `glob` other than `*`/`?` is escaped before compiling, so a name containing e.g. `[`, `(`, or
 * `+` (all valid in a filename) matches literally instead of throwing `SyntaxError` or silently
 * mismatching — this is called with model-suggested glob arguments, so untrusted-shaped input is
 * the expected case, not the exception.
 */
export async function find(fs: DocFs, path: string, glob?: string): Promise<string[]> {
  const re = glob
    ? new RegExp(
        '^' +
          glob.replace(/[.*+?^${}()|[\]\\]/g, (ch) =>
            ch === '*' ? '.*' : ch === '?' ? '.' : `\\${ch}`,
          ) +
          '$',
      )
    : undefined;
  const out: string[] = [];
  async function walk(dir: string): Promise<void> {
    for (const e of await fs.readdir(dir)) {
      const child = `${dir === '/' ? '' : dir}/${e.name}`;
      if (e.kind === 'dir') await walk(child);
      else if (!re || re.test(e.name)) out.push(child);
    }
  }
  await walk(path);
  return out;
}
