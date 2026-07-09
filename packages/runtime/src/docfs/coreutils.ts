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

/**
 * The file-level `tail` coreutil — the LAST `n` lines of a DocFs file (default 10). Distinct from
 * `compose.ts`'s pipeline `tail` transform (`(... | tail 5)`, operating on already-materialized
 * `Value` rows); this one is a top-level DocFs read over a saved artifact or document entry, and
 * mirrors `head`'s exact null-handling/splitting semantics: throw on a missing file (same as
 * `head`/`cat`/`wc`, all of which throw rather than degrade to an empty result), and split on
 * `\n` with no trailing-newline trimming, so `head`/`tail` agree on line-splitting for the same
 * file content.
 */
export async function tail(fs: DocFs, path: string, n = 10): Promise<{ lines: string[] }> {
  const v = await fs.readFile(path);
  if (!v) throw new Error(`tail: no such file: ${path}`);
  // `slice(-0) === slice(0)` (the whole array) — guard n<=0 explicitly so "the last 0 lines"
  // means none, not everything.
  if (n <= 0) return { lines: [] };
  return { lines: v.text.split('\n').slice(-n) };
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
