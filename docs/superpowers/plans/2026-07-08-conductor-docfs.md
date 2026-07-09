# Conductor — Phase 1: the DocFs narrow waist — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Introduce a single read/search filesystem abstraction (`DocFs`) over the live document, the workspace, and (later) sources/skills/peers, plus pure `coreutils` over it — the narrow waist the conductor's pull-based context and orchestration build on. Writes stay OUT of the fs (they go through the existing capability/actuation path).

**Architecture:** One `DocFs` interface (in `@ge/contracts`) → a `DocFsRouter` (in `@ge/runtime`) that dispatches a path's first segment to a `Mount`. Phase 1 ships two real mounts — `DocMount` (backed by the existing `DocBridge`) and `WorkMount` (backed by the existing `WorkspaceStore`) — a read-only `Mount` wrapper, and `coreutils` (`ls/cat/grep/head/find/wc`) as pure functions over `DocFs`. Nothing existing is deleted: `WorkspaceStore` and `DocBridge.captureDocState/searchDocument/readAddress` are *reshaped behind* the waist, not replaced.

**Tech Stack:** TypeScript, Zod (contracts), Vitest. Bun workspaces + TS project references. No new runtime deps.

## Global Constraints

- `bun run typecheck` clean · `bun run test` green · `bun run lint` clean (copied from CLAUDE.md "Definition of done").
- **Surface-agnostic:** `runtime`/`contracts` must contain NO Word/Excel/Office.js code; surface specifics stay in `bridge-*`. `DocMount` talks only to the `DocBridge` interface.
- **Reads/search only in `DocFs`.** Mutation is never a `DocFs` operation; it stays on `DocBridge.actuate` (the gated, provenanced path). Non-`/work` mounts are read-only.
- **Untrusted host content:** everything `DocFs` returns is data, never instructions (ADR-0003).
- **Budgeted:** every `readFile`/`search` is bounded (byte/line/match caps) — mirror the existing `WorkspaceStore` caps (`MAX_ARTIFACT_BYTES`, `PREVIEW_LINES`, `GREP_MATCH_CAP`).
- Paths are POSIX-style, absolute, `/<mount>/<rest>`; mounts are `doc`, `work`, `sources`, `skills`, `peers` (Phase 1 implements `doc` + `work`).

## Roadmap (this plan = Phase 1)

1. **DocFs waist + coreutils + Doc/Work mounts** ← THIS PLAN
2. Pull-context: feed the model `readdir('/doc')` + on-demand reads; `count_tokens` + `cache_control`; fold `@ge/content/chunk` into a `readFile` strategy.
3. Grounding mounts: `/sources` (from `listEngineDataStores`) + `/skills` (from `listAvailableAgentViews`); `@`-picker → `reconcile(plan.ground)` → `dataStoreSpecs`.
4. Orchestration loop: planner→commander via `mention://`, gate, actuation, provenance (formalize existing).
5. A2A specialists: read-only `/peers/*` mounts + `delegate` to Agent Engine agents.

---

## File Structure

- Create `packages/contracts/src/docfs.ts` — `DocFs` interface + entry/stat/view/match types + `parseDocPath` util (Zod for the wire-facing result types).
- Create `packages/contracts/src/docfs.test.ts`.
- Create `packages/runtime/src/docfs/mount.ts` — `Mount` interface + `readOnly()` wrapper.
- Create `packages/runtime/src/docfs/router.ts` — `DocFsRouter implements DocFs`.
- Create `packages/runtime/src/docfs/work-mount.ts` — `WorkMount` over `WorkspaceStore`.
- Create `packages/runtime/src/docfs/doc-mount.ts` — `DocMount` over `DocBridge`.
- Create `packages/runtime/src/docfs/coreutils.ts` — `ls/cat/grep/head/find/wc`.
- Create `packages/runtime/src/docfs/index.ts` — `createDocFs(...)` factory + re-exports.
- Tests: `router.test.ts`, `work-mount.test.ts`, `doc-mount.test.ts`, `coreutils.test.ts`, `readonly.test.ts` under `packages/runtime/src/docfs/`.
- Modify `packages/contracts/src/index.ts` (export `./docfs.js`) and `packages/runtime/src/index.ts` (export `./docfs/index.js`).

---

## Task 1: DocFs contract types + path parsing

**Files:**
- Create: `packages/contracts/src/docfs.ts`
- Test: `packages/contracts/src/docfs.test.ts`
- Modify: `packages/contracts/src/index.ts`

**Interfaces:**
- Produces: `DocFs`, `DirEntry`, `FileStat`, `FileView`, `SearchMatch`, `parseDocPath(path): { mount: string; rel: string }`.

- [ ] **Step 1: Write the failing test**

```ts
// packages/contracts/src/docfs.test.ts
import { describe, expect, it } from 'vitest';
import { parseDocPath } from './docfs.js';

describe('parseDocPath', () => {
  it('splits /<mount>/<rest>', () => {
    expect(parseDocPath('/doc/sheets/Q3.tsv')).toEqual({ mount: 'doc', rel: 'sheets/Q3.tsv' });
    expect(parseDocPath('/work')).toEqual({ mount: 'work', rel: '' });
    expect(parseDocPath('/work/')).toEqual({ mount: 'work', rel: '' });
  });
  it('normalizes and rejects traversal + non-absolute', () => {
    expect(parseDocPath('/doc/a/../b')).toEqual({ mount: 'doc', rel: 'b' });
    expect(() => parseDocPath('doc/x')).toThrow(/absolute/);
    expect(() => parseDocPath('/doc/../../etc')).toThrow(/escape/);
    expect(() => parseDocPath('/')).toThrow(/mount/);
  });
});
```

- [ ] **Step 2: Run it, verify it fails**

Run: `npx vitest run packages/contracts/src/docfs.test.ts`
Expected: FAIL, "parseDocPath is not a function".

- [ ] **Step 3: Implement**

```ts
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
```

- [ ] **Step 4: Run tests, verify pass**

Run: `npx vitest run packages/contracts/src/docfs.test.ts` → PASS.

- [ ] **Step 5: Export + commit**

Add to `packages/contracts/src/index.ts`: `export * from './docfs.js';`
```bash
git add packages/contracts/src/docfs.ts packages/contracts/src/docfs.test.ts packages/contracts/src/index.ts
git commit -m "feat(contracts): add DocFs interface + parseDocPath (conductor waist)"
```

---

## Task 2: Mount interface + read-only wrapper

**Files:**
- Create: `packages/runtime/src/docfs/mount.ts`
- Test: `packages/runtime/src/docfs/readonly.test.ts`

**Interfaces:**
- Consumes: `DirEntry`, `FileStat`, `FileView`, `SearchMatch`, `ReadOpts`, `SearchOpts` from `@ge/contracts`.
- Produces: `Mount` (prefix + the four read ops, all rel-path based), `readOnly(mount): Mount` (identity for reads; marker `readonly = true`).

- [ ] **Step 1: Write the failing test**

```ts
// packages/runtime/src/docfs/readonly.test.ts
import { describe, expect, it } from 'vitest';
import { readOnly, type Mount } from './mount.js';

const base: Mount = {
  prefix: 'x',
  readdir: async () => [{ name: 'a', kind: 'file' }],
  stat: async () => null,
  readFile: async () => null,
  search: async () => [],
};

describe('readOnly', () => {
  it('passes reads through and flags readonly', async () => {
    const ro = readOnly(base);
    expect(ro.readonly).toBe(true);
    expect(ro.prefix).toBe('x');
    expect(await ro.readdir('')).toEqual([{ name: 'a', kind: 'file' }]);
  });
});
```

- [ ] **Step 2: Run, verify fail** — `npx vitest run packages/runtime/src/docfs/readonly.test.ts` → FAIL (no module).

- [ ] **Step 3: Implement**

```ts
// packages/runtime/src/docfs/mount.ts
import type { DirEntry, FileStat, FileView, ReadOpts, SearchMatch, SearchOpts } from '@ge/contracts';

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
```

- [ ] **Step 4: Run, verify pass.**

- [ ] **Step 5: Commit**
```bash
git add packages/runtime/src/docfs/mount.ts packages/runtime/src/docfs/readonly.test.ts
git commit -m "feat(runtime): DocFs Mount interface + readOnly wrapper"
```

---

## Task 3: DocFsRouter

**Files:**
- Create: `packages/runtime/src/docfs/router.ts`
- Test: `packages/runtime/src/docfs/router.test.ts`

**Interfaces:**
- Consumes: `DocFs`, `parseDocPath` (`@ge/contracts`); `Mount` (Task 2).
- Produces: `class DocFsRouter implements DocFs` — ctor `(mounts: Mount[])`; dispatches by mount prefix; `stat('/')`/`readdir('/')` list the mount roots; unknown mount → for `readdir`/`search` returns empty? No — throws a clear error; `readFile`/`stat` on a missing path → `null`.

- [ ] **Step 1: Write the failing test**

```ts
// packages/runtime/src/docfs/router.test.ts
import { describe, expect, it } from 'vitest';
import { DocFsRouter } from './router.js';
import type { Mount } from './mount.js';

const work: Mount = {
  prefix: 'work',
  readdir: async (rel) => (rel === '' ? [{ name: 'notes.md', kind: 'file', size: 12 }] : []),
  stat: async (rel) => (rel === 'notes.md' ? { path: '', kind: 'file', size: 12 } : null),
  readFile: async (rel) => (rel === 'notes.md' ? { path: '', text: 'hello world\n', bytes: 12, truncated: false } : null),
  search: async (_rel, p) => (p === 'hello' ? [{ path: '', line: 1, text: 'hello world' }] : []),
};

describe('DocFsRouter', () => {
  const fs = new DocFsRouter([work]);
  it('lists mount roots at /', async () => {
    expect((await fs.readdir('/')).map((e) => e.name)).toContain('work');
  });
  it('dispatches to the mount and stamps the absolute path', async () => {
    expect((await fs.readdir('/work')).map((e) => e.name)).toEqual(['notes.md']);
    const v = await fs.readFile('/work/notes.md');
    expect(v?.text).toBe('hello world\n');
    expect(v?.path).toBe('/work/notes.md');
    const m = await fs.search('/work', 'hello');
    expect(m).toEqual([{ path: '/work/notes.md', line: 1, text: 'hello world' }]);
  });
  it('throws on an unknown mount', async () => {
    await expect(fs.readdir('/nope')).rejects.toThrow(/unknown mount/i);
  });
  it('returns null reading a missing file', async () => {
    expect(await fs.readFile('/work/missing')).toBeNull();
  });
});
```

- [ ] **Step 2: Run, verify fail.**

- [ ] **Step 3: Implement**

```ts
// packages/runtime/src/docfs/router.ts
import { parseDocPath, type DirEntry, type DocFs, type FileStat, type FileView, type ReadOpts, type SearchMatch, type SearchOpts } from '@ge/contracts';
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
```

- [ ] **Step 4: Run tests → PASS.** (Note the test's mount returns `path: ''` for the single file; the router stamps `/work` + rel. For `search`, the mount returns `path: ''` meaning "the searched dir's file" — adjust the WorkMount in Task 5 to return the file's rel path; here the test asserts `/work/notes.md` because rel is `notes.md`. Fix the fixture to return `path: 'notes.md'` if needed and re-run.)

- [ ] **Step 5: Commit**
```bash
git add packages/runtime/src/docfs/router.ts packages/runtime/src/docfs/router.test.ts
git commit -m "feat(runtime): DocFsRouter — dispatch DocFs to mounts by prefix"
```

---

## Task 4: WorkMount (over the existing WorkspaceStore)

**Files:**
- Create: `packages/runtime/src/docfs/work-mount.ts`
- Test: `packages/runtime/src/docfs/work-mount.test.ts`
- Reference: `packages/runtime/src/workspace.ts` (`WorkspaceStore`: `list/summary/save/cat/grep`, `WorkspaceArtifact{ id,name,bytes,lineCount,text }`).

**Interfaces:**
- Consumes: `WorkspaceStore` (existing), `Mount` (Task 2).
- Produces: `workMount(store: WorkspaceStore): Mount` — `prefix:'work'`; `readdir('')` → one `DirEntry` per artifact (`name`, `size: bytes`); `readFile(name)` → the artifact text (respect `maxBytes`); `search(rel, pattern)` → grep across artifacts (or one when `rel` names an artifact); `stat(name)` → size/kind.

- [ ] **Step 1: Write the failing test**

```ts
// packages/runtime/src/docfs/work-mount.test.ts
import { describe, expect, it } from 'vitest';
import { WorkspaceStore } from '../workspace.js';
import { workMount } from './work-mount.js';

function seeded() {
  const s = new WorkspaceStore();
  s.save({ name: 'q3.tsv', sourceLabel: 'test', content: 'a\tb\n1\t2\n', kind: 'tsv' });
  return s;
}

describe('workMount', () => {
  it('lists artifacts as files', async () => {
    const m = workMount(seeded());
    const entries = await m.readdir('');
    expect(entries.map((e) => e.name)).toContain('q3.tsv');
    expect(entries[0]!.kind).toBe('file');
  });
  it('reads an artifact by name', async () => {
    const m = workMount(seeded());
    const v = await m.readFile('q3.tsv');
    expect(v?.text).toContain('a\tb');
  });
  it('greps across artifacts and returns the artifact-relative path', async () => {
    const m = workMount(seeded());
    const hits = await m.search('', '1');
    expect(hits[0]).toMatchObject({ path: 'q3.tsv', line: 2 });
  });
  it('returns null for a missing artifact', async () => {
    expect(await workMount(seeded()).readFile('nope')).toBeNull();
  });
});
```

- [ ] **Step 2: Run, verify fail.**

- [ ] **Step 3: Implement** (adapt to the real `WorkspaceStore` method names/return shapes; read `workspace.ts` first)

```ts
// packages/runtime/src/docfs/work-mount.ts
import type { DirEntry, FileView, ReadOpts, SearchMatch, SearchOpts } from '@ge/contracts';
import type { WorkspaceStore } from '../workspace.js';
import type { Mount } from './mount.js';

const DEFAULT_MAX_BYTES = 256 * 1024;

/** `/work` — the in-memory WorkspaceStore artifacts as files. Reshape, not rewrite: this delegates
 *  to the store's existing list/cat/grep so there is one artifact store, behind the DocFs waist. */
export function workMount(store: WorkspaceStore): Mount {
  const list = () => store.run({ workspace: 'list' }); // adapt to the real accessor; see workspace.ts
  return {
    prefix: 'work',
    async readdir() {
      const res = list();
      return res.workspace === 'list'
        ? res.artifacts.map((a): DirEntry => ({ name: a.name, kind: 'file', size: a.bytes }))
        : [];
    },
    async stat(rel) {
      const res = list();
      const a = res.workspace === 'list' ? res.artifacts.find((x) => x.name === rel) : undefined;
      return a ? { path: '', kind: 'file', size: a.bytes } : null;
    },
    async readFile(rel, opts?: ReadOpts): Promise<FileView | null> {
      const res = store.run({ workspace: 'cat', name: rel, head: Number.MAX_SAFE_INTEGER });
      if (res.workspace !== 'cat') return null;
      const max = opts?.maxBytes ?? DEFAULT_MAX_BYTES;
      const text = res.preview.slice(0, max);
      return { path: '', text, bytes: Buffer.byteLength(text), truncated: text.length < res.preview.length };
    },
    async search(rel, pattern, opts?: SearchOpts): Promise<SearchMatch[]> {
      const res = store.run({ workspace: 'grep', name: rel || undefined, pattern, context: opts?.context ?? 0 });
      if (res.workspace !== 'grep') return [];
      return res.matches.slice(0, opts?.max ?? 50).map((m) => ({ path: res.artifact.name, line: m.line, text: m.text }));
    },
  };
}
```

> **Note for the implementer:** `packages/runtime/src/workspace.ts` exposes `WorkspaceResult` variants (`list/summary/save/cat/grep`). Read it and wire `workMount` to the actual accessor (e.g. a `run(input)` dispatcher or discrete methods). Keep the caps the store already enforces; do not re-implement grep.

- [ ] **Step 4: Run tests → PASS** (fix method wiring against the real `WorkspaceStore` API).

- [ ] **Step 5: Commit**
```bash
git add packages/runtime/src/docfs/work-mount.ts packages/runtime/src/docfs/work-mount.test.ts
git commit -m "feat(runtime): WorkMount — WorkspaceStore behind the DocFs waist"
```

---

## Task 5: DocMount (over the existing DocBridge)

**Files:**
- Create: `packages/runtime/src/docfs/doc-mount.ts`
- Test: `packages/runtime/src/docfs/doc-mount.test.ts`
- Reference: `packages/runtime/src/bridge.ts` (`DocBridge.captureDocState`, `searchDocument`, `readAddress`), `packages/contracts/src/doc-state.ts` (`DocStateSnapshot` outline/inventory/selection/namedRanges/comments).

**Interfaces:**
- Consumes: `DocBridge` (existing), `Mount`.
- Produces: `docMount(bridge: DocBridge): Mount` — read-only-by-nature:
  - `readdir('')` → the doc-state **inventory + outline** as entries (e.g. `sheets/`, `outline.md`, `selection`, `named-ranges.json`); a tiny, structured index — the pull-context "table of contents".
  - `readFile('selection')` / `readFile('outline.md')` / `readFile('<inventory-id>')` → lazy content via `readAddress`/`searchDocument`, budgeted.
  - `search(rel, pattern)` → `bridge.searchDocument(pattern)` mapped to `SearchMatch[]`.
  - `stat` → from the inventory.

- [ ] **Step 1: Write the failing test** (with a fake `DocBridge`)

```ts
// packages/runtime/src/docfs/doc-mount.test.ts
import { describe, expect, it } from 'vitest';
import type { DocBridge } from '../bridge.js';
import { docMount } from './doc-mount.js';

const bridge = {
  surface: 'excel',
  async captureDocState() {
    return {
      surface: 'excel',
      outline: [{ level: 1, text: 'Summary' }],
      inventory: [{ kind: 'sheet', id: 'sheet:Q3', title: 'Q3', summary: '9x4' }],
      selection: { kind: 'range', title: 'A1:D9', preview: 'rev…' },
    };
  },
  async searchDocument(q: string) {
    return q === 'revenue' ? [{ kind: 'range', id: 'r1', value: { text: 'revenue 42' } }] : [];
  },
} as unknown as DocBridge;

describe('docMount', () => {
  it('lists the doc-state index (inventory + selection + outline)', async () => {
    const entries = await docMount(bridge).readdir('');
    const names = entries.map((e) => e.name);
    expect(names).toContain('outline.md');
    expect(names).toContain('selection');
    expect(names).toContain('sheet:Q3');
  });
  it('reads the outline', async () => {
    const v = await docMount(bridge).readFile('outline.md');
    expect(v?.text).toContain('Summary');
  });
  it('search delegates to searchDocument', async () => {
    const hits = await docMount(bridge).search('', 'revenue');
    expect(hits[0]?.text).toContain('revenue 42');
  });
});
```

- [ ] **Step 2: Run, verify fail.**

- [ ] **Step 3: Implement** (map `DocStateSnapshot` → entries; degrade gracefully when optional bridge methods are absent)

```ts
// packages/runtime/src/docfs/doc-mount.ts
import type { DirEntry, FileView, ReadOpts, SearchMatch, SearchOpts } from '@ge/contracts';
import type { DocBridge } from '../bridge.js';
import type { Mount } from './mount.js';

const MAX = 64 * 1024;

/** `/doc` — the live document as a lazy, read-only index: readdir = the doc-state inventory/outline,
 *  readFile/search = on-demand reads via the bridge. Surface-agnostic (DocBridge only). */
export function docMount(bridge: DocBridge): Mount {
  return {
    prefix: 'doc',
    readonly: true,
    async readdir() {
      const s = (await bridge.captureDocState?.()) ?? undefined;
      const entries: DirEntry[] = [{ name: 'outline.md', kind: 'file' }];
      if (s?.selection) entries.push({ name: 'selection', kind: 'file' });
      if (s?.namedRanges?.length) entries.push({ name: 'named-ranges.json', kind: 'file' });
      for (const inv of s?.inventory ?? []) entries.push({ name: inv.id, kind: 'file', size: undefined });
      return entries;
    },
    async stat(rel) {
      const entries = await this.readdir('');
      const e = entries.find((x) => x.name === rel);
      return e ? { path: '', kind: 'file', size: e.size ?? 0 } : null;
    },
    async readFile(rel, opts?: ReadOpts): Promise<FileView | null> {
      const s = (await bridge.captureDocState?.()) ?? undefined;
      const cap = (t: string) => {
        const text = t.slice(0, opts?.maxBytes ?? MAX);
        return { path: '', text, bytes: Buffer.byteLength(text), truncated: text.length < t.length };
      };
      if (rel === 'outline.md' && s)
        return cap(s.outline.map((o) => `${'#'.repeat(Math.max(1, o.level))} ${o.text}`).join('\n'));
      if (rel === 'selection' && s?.selection) return cap(s.selection.preview ?? s.selection.title);
      if (rel === 'named-ranges.json' && s?.namedRanges) return cap(JSON.stringify(s.namedRanges, null, 2));
      // inventory entry → lazy address read when the bridge supports it
      const parts = (await bridge.readAddress?.(rel)) ?? [];
      const text = parts.map((p) => renderResolved(p)).join('\n');
      return text ? cap(text) : null;
    },
    async search(_rel, pattern, opts?: SearchOpts): Promise<SearchMatch[]> {
      const parts = (await bridge.searchDocument?.(pattern)) ?? [];
      return parts
        .slice(0, opts?.max ?? 50)
        .map((p, i): SearchMatch => ({ path: (p as { id?: string }).id ?? `hit:${i}`, line: 1, text: renderResolved(p) }));
    },
  };
}

/** Render a ResolvedContext value to a text line for DocFs views (data only, never instructions). */
function renderResolved(part: unknown): string {
  const v = (part as { value?: { text?: string } }).value;
  return v?.text ?? JSON.stringify(part).slice(0, 500);
}
```

> **Implementer note:** confirm `ResolvedContext.value` shape in `contracts/context.ts` and adjust `renderResolved`. `readAddress`/`searchDocument`/`captureDocState` are all OPTIONAL on `DocBridge` — the `?.()` guards keep DocMount working on surfaces that omit them (degrade to empty/outline-only).

- [ ] **Step 4: Run tests → PASS.**

- [ ] **Step 5: Commit**
```bash
git add packages/runtime/src/docfs/doc-mount.ts packages/runtime/src/docfs/doc-mount.test.ts
git commit -m "feat(runtime): DocMount — live document as a lazy read-only DocFs index"
```

---

## Task 6: coreutils over DocFs (pure functions)

**Files:**
- Create: `packages/runtime/src/docfs/coreutils.ts`
- Test: `packages/runtime/src/docfs/coreutils.test.ts`

**Interfaces:**
- Consumes: `DocFs`.
- Produces: `ls(fs, path)`, `cat(fs, path, opts?)`, `grep(fs, path, pattern, opts?)`, `head(fs, path, n?)`, `find(fs, path, glob?)`, `wc(fs, path)` — each returns a small structured result AND `.toString()`-able text. Pure over the `DocFs` interface (no host coupling).

- [ ] **Step 1: Write the failing test** (against `DocFsRouter` + `workMount` from earlier tasks)

```ts
// packages/runtime/src/docfs/coreutils.test.ts
import { describe, expect, it } from 'vitest';
import { WorkspaceStore } from '../workspace.js';
import { DocFsRouter } from './router.js';
import { workMount } from './work-mount.js';
import { cat, grep, head, ls, wc } from './coreutils.js';

function fs() {
  const s = new WorkspaceStore();
  s.save({ name: 'a.txt', sourceLabel: 't', content: 'one\ntwo\nthree\n' });
  return new DocFsRouter([workMount(s)]);
}

describe('coreutils', () => {
  it('ls lists a directory', async () => expect(await ls(fs(), '/work')).toEqual(['a.txt']));
  it('cat returns file text', async () => expect((await cat(fs(), '/work/a.txt')).text).toContain('two'));
  it('head returns first n lines', async () => expect((await head(fs(), '/work/a.txt', 2)).lines).toEqual(['one', 'two']));
  it('grep finds matches with line numbers', async () => {
    const r = await grep(fs(), '/work', 'two');
    expect(r.matches[0]).toMatchObject({ path: '/work/a.txt', line: 2, text: 'two' });
  });
  it('wc counts lines/bytes', async () => expect((await wc(fs(), '/work/a.txt')).lines).toBe(3));
});
```

- [ ] **Step 2: Run, verify fail.**

- [ ] **Step 3: Implement**

```ts
// packages/runtime/src/docfs/coreutils.ts
import type { DocFs, SearchMatch } from '@ge/contracts';

export async function ls(fs: DocFs, path: string): Promise<string[]> {
  return (await fs.readdir(path)).map((e) => (e.kind === 'dir' ? `${e.name}/` : e.name));
}

export async function cat(fs: DocFs, path: string, opts?: { maxBytes?: number }): Promise<{ text: string; truncated: boolean }> {
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
  return { lines: v.text.split('\n').filter((l, i, a) => i < a.length - 1 || l.length > 0).length, bytes: v.bytes };
}

/** Recursive listing with an optional `*`/`?` glob on the leaf name. */
export async function find(fs: DocFs, path: string, glob?: string): Promise<string[]> {
  const re = glob ? new RegExp('^' + glob.replace(/[.]/g, '\\.').replace(/\*/g, '.*').replace(/\?/g, '.') + '$') : undefined;
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
```

- [ ] **Step 4: Run tests → PASS.**

- [ ] **Step 5: Commit**
```bash
git add packages/runtime/src/docfs/coreutils.ts packages/runtime/src/docfs/coreutils.test.ts
git commit -m "feat(runtime): coreutils (ls/cat/grep/head/find/wc) over DocFs"
```

---

## Task 7: `createDocFs` factory + exports + integration test

**Files:**
- Create: `packages/runtime/src/docfs/index.ts`
- Test: `packages/runtime/src/docfs/index.test.ts`
- Modify: `packages/runtime/src/index.ts`

**Interfaces:**
- Consumes: `DocFsRouter`, `docMount`, `workMount`, `readOnly` (earlier tasks), `DocBridge`, `WorkspaceStore`.
- Produces: `createDocFs(opts: { bridge: DocBridge; workspace: WorkspaceStore }): DocFs` — assembles `[readOnly(docMount(bridge)), workMount(workspace)]` into a `DocFsRouter`; re-exports the DocFs pieces + coreutils.

- [ ] **Step 1: Write the failing test**

```ts
// packages/runtime/src/docfs/index.test.ts
import { describe, expect, it } from 'vitest';
import { WorkspaceStore } from '../workspace.js';
import type { DocBridge } from '../bridge.js';
import { createDocFs, ls } from './index.js';

const bridge = { surface: 'excel', async captureDocState() { return { surface: 'excel', outline: [], inventory: [] }; } } as unknown as DocBridge;

describe('createDocFs', () => {
  it('mounts /doc (read-only) and /work', async () => {
    const fs = createDocFs({ bridge, workspace: new WorkspaceStore() });
    const roots = await ls(fs, '/');
    expect(roots).toEqual(['doc/', 'work/']);
  });
});
```

- [ ] **Step 2: Run, verify fail.**

- [ ] **Step 3: Implement**

```ts
// packages/runtime/src/docfs/index.ts
import type { DocFs } from '@ge/contracts';
import type { DocBridge } from '../bridge.js';
import type { WorkspaceStore } from '../workspace.js';
import { DocFsRouter } from './router.js';
import { docMount } from './doc-mount.js';
import { workMount } from './work-mount.js';
import { readOnly } from './mount.js';

export * from './mount.js';
export * from './router.js';
export * from './coreutils.js';
export { docMount } from './doc-mount.js';
export { workMount } from './work-mount.js';

/** Assemble the Phase-1 DocFs: `/doc` (read-only, live document) + `/work` (WorkspaceStore). */
export function createDocFs(opts: { bridge: DocBridge; workspace: WorkspaceStore }): DocFs {
  return new DocFsRouter([readOnly(docMount(opts.bridge)), workMount(opts.workspace)]);
}
```

- [ ] **Step 4: Run tests → PASS.**

- [ ] **Step 5: Export + full gate + commit**

Add to `packages/runtime/src/index.ts`: `export * from './docfs/index.js';`
```bash
bun run typecheck && npx vitest run packages/runtime/src/docfs packages/contracts/src/docfs.test.ts && npx prettier --check "packages/**/src/docfs/**" "packages/contracts/src/docfs.ts"
git add packages/runtime/src/docfs packages/runtime/src/index.ts
git commit -m "feat(runtime): createDocFs factory + exports (conductor Phase 1 complete)"
```

---

## Self-Review

- **Spec coverage:** Phase-1 scope = the waist (Task 1), Mount+readOnly (Task 2), router (Task 3), the two real mounts (Tasks 4–5), coreutils (Task 6), factory+wiring (Task 7). `/sources`, `/skills`, `/peers`, pull-context, and orchestration are explicitly Phases 2–5, out of scope here. ✔
- **Placeholder scan:** the two "implementer note" blocks (WorkMount accessor names, `ResolvedContext.value` shape) are the only "verify against real API" spots — flagged deliberately because those exact shapes must be read from `workspace.ts`/`context.ts` at implementation time; every step ships real code. Fix the fixture path detail called out in Task 3 Step 4.
- **Type consistency:** `Mount` ops take mount-relative paths and return mount-relative `path` fields; `DocFsRouter` stamps them absolute — consistent across Tasks 2/3/4/5. `DocFs` signatures identical in `contracts` and every consumer.
- **Constraints:** reads-only interface (writes stay on `actuate`), surface-agnostic (`DocMount` uses only `DocBridge`), budgeted (byte/line/match caps), no new deps. ✔
