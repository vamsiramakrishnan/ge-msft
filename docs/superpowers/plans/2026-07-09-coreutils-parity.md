# Coreutils Parity (`sed`/`derive`, `tail`, `cp`/`mv`/`rm`, `/work` namespacing) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the highest-value gaps between our DocFs/pipeline coreutils and Anthropic's conductor's `ls/cat/grep/sed/awk/head/tail/wc/find/glob/cp/mv/rm/mkdir` set — scoped by three explicit decisions the repo owner made up front (see Global Constraints).

**Architecture:** Two separate, already-existing subsystems each gain capability, rather than one new command surface:
1. The **pure pipeline-transform registry** (`packages/runtime/src/compose.ts` + `TRANSFORM_NAMES` in `packages/contracts/src/expr-grammar.ts`) — a `Transform = (input: Value, rawArgs: string) => Value | EvalError` over `Value = table | number | text`, composed via `save x = (read ... | transform ... | transform ...)`. Gains `sed` (text/cell substitution) and `derive` (arithmetic computed column) — the sed/awk-equivalent capability, as pure functions slotting into existing, tested infrastructure.
2. The **DocFs coreutils + WorkspaceStore** (`packages/runtime/src/docfs/coreutils.ts`, `packages/runtime/src/workspace.ts`) — gains a `tail` file-level coreutil (mirrors the existing `head`, wired as a new top-level read verb exactly like `ls`/`find` were in the prior plan) and `cp`/`mv`/`rm` artifact-lifecycle operations on `/work` (local scratch state only — `/doc`, the live document, stays strictly read-only per `docfs.ts`'s own contract; nothing here touches the gated `DocBridge.actuate` path).

**Tech Stack:** Same as the rest of the repo — TypeScript, Zod, Vitest, Bun workspaces; Python `parse_commands.py` for skill-side parity where new top-level verbs are added (transforms validate against a manifest-driven set already, per Task 6's investigation step).

## Global Constraints

- **`sed`/`derive` land as pipeline transforms, not new top-level commands** (repo owner's explicit choice) — they only exist inside a `(... | sed ... | ...)` composition, never as a bare `sed <path>` command.
- **The existing `cat`/`grep` workspace verbs are NOT touched in this plan** (repo owner's explicit choice, respecting this repo's own prior deferral of the cat/grep-over-DocFs unification, pending the incident investigation's capture-legibility findings). Do not rename, extend, or reroute them. Do not add a standalone `glob` verb either — `find`'s existing optional glob argument already covers pattern-based discovery, and true multi-artifact batch operations (the other half of `glob`'s value) wait on that same deferred unification.
- **No `mkdir` verb, no new stateful "directory" entity** (repo owner's explicit choice). `/` inside a saved artifact name is a naming convention only, rendered as a pseudo-directory purely at `readdir`/`find` time over `/work` — no new persisted state, no orphaned-empty-directory case to handle.
- **`cp`/`mv`/`rm` operate ONLY on `/work` (WorkspaceStore) artifacts.** They must not be reachable against `/doc` — `docMount` stays wrapped in `readOnly(...)` exactly as today; do not touch that wrapping.
- Follow this repo's `ls`/`find` wiring precedent exactly for every new top-level verb: `command-grammar.ts` (`READ_VERBS`/union/schema/`parseCommandLine`/`grammarFor`) → `capability.ts` (`ReadVerbSchema` if applicable) → `command-protocol.ts` (`ReadIntent`/`compileCommand`) → `assist-session.ts` (`READ_COMMAND_VERBS` set + `runReadIntent` case) → `command-help.ts` → regenerate the skill catalog → Python `parse_commands.py` parity. Tasks 1-5 of the prior `docfs-loop-skill-wiring` plan hit type-checking gaps at several of these exact seams (a `ParsedCommandSchema` branch missing is invisible to `tsc` because of covariant Zod-schema typing; `capability.ts`'s `ReadVerbSchema` enum has no drift guard against `READ_VERBS`; `assist-session.ts`'s `READ_COMMAND_VERBS` is a THIRD hand-maintained verb set) — expect the same seams here and check each one explicitly per task, don't assume `tsc` alone will catch a missing branch.
- Any gate command a subagent runs (`bun run typecheck`, `npx vitest run ...`, `python3 ...`) MUST be wrapped in `timeout Ns` — this repo has had subagent Bash calls hang silently for 45+ minutes with zero live process. A timeout turns a hang into a fast, visible failure.
- Mixed commits that absorb pre-existing, unrelated, already-in-flight uncommitted WIP are explicitly acceptable in this repo (confirmed policy) — do not treat commit scope/size/staging purity as a defect in any task review.

---

### Task 1: `sed` pipeline transform

**Files:**
- Modify: `packages/contracts/src/expr-grammar.ts` (`TRANSFORM_NAMES`)
- Modify: `packages/runtime/src/compose.ts` (`TRANSFORMS`, `TRANSFORM_USAGE`, new `sed` function)
- Test: `packages/runtime/src/compose.test.ts` (find the existing transform tests — e.g. `filter`/`select` — and mirror their pattern)

**Interfaces:**
- Consumes: `Value` (`{kind:'table', columns, rows}` | `{kind:'number', value}` | `{kind:'text', value}`), `Transform = (input: Value, rawArgs: string) => Value | EvalError` (both already defined in `compose.ts`).
- Produces: `sed(input: Value, rawArgs: string): Value | EvalError` registered in `TRANSFORMS.sed`, added to `TRANSFORM_NAMES` and `TRANSFORM_USAGE`.

- [ ] **Step 1: Read the existing transform pattern**

Read `packages/runtime/src/compose.ts`'s `filter` and `select` functions in full (search `function filter`, `function select`), and the `requireTable` helper. Read `packages/runtime/src/compose.test.ts` for how an existing transform (e.g. `filter` or `sort`) is unit-tested directly (not through the full `save`/pipeline parser) — mirror that exact test-construction pattern, not a new one.

- [ ] **Step 2: Write the failing tests**

```ts
it('sed replaces the first match per row cell on a table (no /g flag)', () => {
  const table: Value = { kind: 'table', columns: ['Region'], rows: [['East Coast'], ['West Coast']] };
  const out = sed(table, "s/Coast/Region/");
  expect(out).toEqual({ kind: 'table', columns: ['Region'], rows: [['East Region'], ['West Region']] });
});

it('sed with /g replaces every match in a cell', () => {
  const table: Value = { kind: 'table', columns: ['Label'], rows: [['aa-aa']] };
  expect(sed(table, 's/a/x/g')).toEqual({ kind: 'table', columns: ['Label'], rows: [['xx-xx']] });
});

it('sed operates on a text Value directly', () => {
  expect(sed({ kind: 'text', value: 'hello world' }, 's/world/there/')).toEqual({
    kind: 'text',
    value: 'hello there',
  });
});

it('sed rejects a number Value', () => {
  const out = sed({ kind: 'number', value: 42 }, 's/4/9/');
  expect(out).toHaveProperty('error');
});

it('sed rejects a malformed s/// expression', () => {
  const out = sed({ kind: 'text', value: 'x' }, 'not-a-sed-expr');
  expect(out).toHaveProperty('error');
  expect((out as { error: string }).error).toMatch(/usage: sed/i);
});
```

- [ ] **Step 3: Run, verify fail**

Run: `timeout 30 npx vitest run packages/runtime/src/compose.test.ts -t "sed"`
Expected: FAIL — `sed` is not exported/defined.

- [ ] **Step 4: Implement `sed`**

In `packages/runtime/src/compose.ts`, add (near `filter`/`select`):

```ts
/** `sed 's/pattern/replacement/[g]'` — literal-or-regex substitution, table cells or text. */
function sed(input: Value, rawArgs: string): Value | EvalError {
  const m = /^s\/((?:[^\\/]|\\.)*)\/((?:[^\\/]|\\.)*)\/(g?)$/.exec(rawArgs.trim());
  if (!m) return { error: 'sed usage: sed s/pattern/replacement/[g] — e.g. sed s/Coast/Region/g' };
  const [, pattern, replacement, flags] = m;
  let re: RegExp;
  try {
    re = new RegExp(pattern!.replace(/\\\//g, '/'), flags === 'g' ? 'g' : '');
  } catch {
    return { error: `sed: invalid pattern — ${pattern}` };
  }
  const repl = replacement!.replace(/\\\//g, '/');
  if (input.kind === 'text') return { kind: 'text', value: input.value.replace(re, repl) };
  if (input.kind === 'table') {
    return {
      kind: 'table',
      columns: input.columns,
      rows: input.rows.map((row) => row.map((cell) => cell.replace(re, repl))),
    };
  }
  return { error: 'sed expects a table or text, got a number' };
}
```

Register it: add `sed,` to `TRANSFORMS`, `'sed',` to `TRANSFORM_NAMES` (in `expr-grammar.ts`), and `sed: 'sed s/pattern/replacement/[g] — text/cell substitution',` to `TRANSFORM_USAGE`.

**Note on regex safety:** this is model-authored input reaching `new RegExp` — the `try/catch` above already turns a malformed pattern into a corrective error rather than a throw. Do not add ReDoS-mitigation beyond that in this task (out of scope) — flag it in your report if you think it's warranted, but do not silently implement additional guards not in this brief.

- [ ] **Step 5: Run, verify pass**

Run: `timeout 30 npx vitest run packages/runtime/src/compose.test.ts -t "sed"`

- [ ] **Step 6: Full gate + commit**

```bash
timeout 120 bun run typecheck
timeout 60 npx vitest run packages/runtime/src/compose.test.ts packages/contracts/src/expr-grammar.test.ts
timeout 30 npx prettier --check packages/runtime/src/compose.ts packages/runtime/src/compose.test.ts packages/contracts/src/expr-grammar.ts
git add packages/runtime/src/compose.ts packages/runtime/src/compose.test.ts packages/contracts/src/expr-grammar.ts
git commit -m "feat(runtime): add sed pipeline transform (text/cell substitution)"
```

---

### Task 2: `derive` pipeline transform (computed column)

**Files:**
- Modify: `packages/contracts/src/expr-grammar.ts` (`TRANSFORM_NAMES`)
- Modify: `packages/runtime/src/compose.ts` (`TRANSFORMS`, `TRANSFORM_USAGE`, new `derive` function)
- Test: `packages/runtime/src/compose.test.ts`

**Interfaces:**
- Consumes: same `Value`/`Transform` as Task 1.
- Produces: `derive(input: Value, rawArgs: string): Value | EvalError` — appends one new numeric column computed from two existing numeric columns (or one column and a literal), via `+`/`-`/`*`/`/`. This is deliberately minimal (no chained expressions, no functions) — the one concrete gap identified versus `awk`: a computed/derived column, not a general expression language.

- [ ] **Step 1: Read the existing `aggregate`/`filter` numeric-parsing pattern**

Read how `filter`/`aggregate` in `compose.ts` already parse a cell as a number (search for `Number(` or `parseFloat` in that file) — reuse the exact same numeric-coercion approach so `derive`'s "not a number" behavior is consistent with the rest of the transform registry.

- [ ] **Step 2: Write the failing tests**

```ts
it('derive appends a computed column from two existing columns', () => {
  const table: Value = {
    kind: 'table',
    columns: ['Budget', 'Actual'],
    rows: [['100', '80'], ['200', '250']],
  };
  const out = derive(table, 'Variance = Budget - Actual');
  expect(out).toEqual({
    kind: 'table',
    columns: ['Budget', 'Actual', 'Variance'],
    rows: [['100', '80', '20'], ['200', '250', '-50']],
  });
});

it('derive supports a column and a literal operand', () => {
  const table: Value = { kind: 'table', columns: ['Revenue'], rows: [['100']] };
  expect(derive(table, 'Doubled = Revenue * 2')).toEqual({
    kind: 'table',
    columns: ['Revenue', 'Doubled'],
    rows: [['200']],
  });
});

it('derive rejects a reference to an unknown column', () => {
  const table: Value = { kind: 'table', columns: ['A'], rows: [['1']] };
  const out = derive(table, 'X = A - B');
  expect(out).toHaveProperty('error');
  expect((out as { error: string }).error).toMatch(/unknown column.*B/i);
});

it('derive rejects a non-numeric cell for the row it fails on', () => {
  const table: Value = { kind: 'table', columns: ['A', 'B'], rows: [['x', '1']] };
  const out = derive(table, 'C = A + B');
  expect(out).toHaveProperty('error');
});

it('derive rejects a non-table Value', () => {
  expect(derive({ kind: 'number', value: 1 }, 'X = A + B')).toHaveProperty('error');
});

it('derive rejects a malformed expression', () => {
  expect(derive({ kind: 'table', columns: ['A'], rows: [] }, 'not an expr')).toHaveProperty('error');
});
```

- [ ] **Step 3: Run, verify fail**

Run: `timeout 30 npx vitest run packages/runtime/src/compose.test.ts -t "derive"`

- [ ] **Step 4: Implement `derive`**

Grammar: `<newCol> = <col|number> <op> <col|number>` where `<op>` is one of `+ - * /`. Implementation sketch (adapt to match the exact numeric-parsing helper found in Step 1 rather than reintroducing a second one):

```ts
const DERIVE_RE = /^(\S+)\s*=\s*(\S+)\s*([+\-*/])\s*(\S+)$/;

function derive(input: Value, rawArgs: string): Value | EvalError {
  const t = requireTable(input, 'derive');
  if ('error' in t) return t;
  const m = DERIVE_RE.exec(rawArgs.trim());
  if (!m) return { error: 'derive usage: derive <newCol> = <col|num> <+|-|*|/> <col|num>' };
  const [, newCol, lhsRaw, op, rhsRaw] = m;
  const { columns, rows } = t.table;
  const resolve = (raw: string, rowIdx: number): number | EvalError => {
    const idx = colIndex(columns, raw!);
    if (idx === -1) {
      const n = Number(raw);
      if (Number.isNaN(n)) return { error: `derive: unknown column or number "${raw}"` };
      return n;
    }
    const n = Number(rows[rowIdx]![idx]);
    if (Number.isNaN(n)) return { error: `derive: non-numeric cell in column "${raw}" (row ${rowIdx + 1})` };
    return n;
  };
  const outRows: string[][] = [];
  for (let i = 0; i < rows.length; i++) {
    const lhs = resolve(lhsRaw!, i);
    if (typeof lhs !== 'number') return lhs;
    const rhs = resolve(rhsRaw!, i);
    if (typeof rhs !== 'number') return rhs;
    const value = op === '+' ? lhs + rhs : op === '-' ? lhs - rhs : op === '*' ? lhs * rhs : lhs / rhs;
    outRows.push([...rows[i]!, String(value)]);
  }
  return { kind: 'table', columns: [...columns, newCol!], rows: outRows };
}
```

Register it in `TRANSFORMS`/`TRANSFORM_NAMES`/`TRANSFORM_USAGE` the same way as Task 1's `sed`.

- [ ] **Step 5: Run, verify pass**
- [ ] **Step 6: Full gate + commit** (same shape as Task 1 Step 6, commit message: `feat(runtime): add derive pipeline transform (computed column)`)

---

### Task 3: `tail` — file-level DocFs coreutil + top-level read verb

**Files:**
- Modify: `packages/runtime/src/docfs/coreutils.ts` (new `tail` function, mirrors `head`)
- Modify: `packages/contracts/src/command-grammar.ts` (`READ_VERBS`, `ParsedCommand` union, `ParsedCommandSchema`, `parseCommandLine`, `grammarFor`)
- Modify: `packages/contracts/src/capability.ts` (`ReadVerbSchema` — check if it needs `'tail'`; per this plan's Global Constraints, this repo's `ls`/`find` work needed this exact addition both times)
- Modify: `packages/runtime/src/command-protocol.ts` (`ReadIntent`, `compileCommand`)
- Modify: `packages/runtime/src/assist-session.ts` (`READ_COMMAND_VERBS`, `runReadIntent` case)
- Modify: `packages/contracts/src/command-help.ts`
- Test: `packages/runtime/src/docfs/coreutils.test.ts`, `packages/contracts/src/command-grammar.test.ts`, `packages/runtime/src/command-protocol.test.ts`, `packages/runtime/src/assist-session.test.ts`

**Interfaces:**
- Consumes: `DocFs` (`readFile`), same as `head`'s existing implementation in `coreutils.ts`.
- Produces: `tail(fs: DocFs, path: string, n?: number): Promise<{ lines: string[] }>`; `ParsedCommand` variant `{ verb: 'tail'; path: string; n?: number }`; `ReadIntent` variant `{ read: 'tail'; path: string; n?: number }`.

**Do NOT confuse this with the existing pipeline `tail` transform** (`compose.ts`'s `takeRows('tail')`, already registered under `TRANSFORM_NAMES` — operates on pipeline `Value` rows, e.g. `(... | tail 5)`). This task adds a DIFFERENT, file-oriented top-level command (`tail /work/schedule.tsv 20` — last N lines of a saved artifact or document entry) in a different grammar slot (bare command, not a pipeline stage). No collision, but name the two clearly in your commit message so a future reader isn't confused.

- [ ] **Step 1: Read the existing `head` coreutil and the full `ls`/`find` wiring**

Read `packages/runtime/src/docfs/coreutils.ts`'s `head` function in full. Read `packages/contracts/src/command-grammar.ts`'s `case 'find':` and its `ParsedCommandSchema` branch, `packages/runtime/src/command-protocol.ts`'s `case 'find':`, and `packages/runtime/src/assist-session.ts`'s `READ_COMMAND_VERBS` set and its `case 'find':` in `runReadIntent` — `tail` mirrors `find`'s shape exactly (one required path, one optional numeric/string argument) more closely than `ls`'s shape (path-only).

- [ ] **Step 2: Write the failing coreutils test**

```ts
it('tail returns the last n lines of a file (default 10)', async () => {
  const fs = fakeFsWithFile('notes.md', Array.from({ length: 15 }, (_, i) => `line ${i + 1}`).join('\n'));
  const { lines } = await tail(fs, '/work/notes.md');
  expect(lines).toEqual(Array.from({ length: 10 }, (_, i) => `line ${i + 6}`));
});

it('tail respects an explicit n', async () => {
  const fs = fakeFsWithFile('notes.md', 'a\nb\nc\nd');
  expect((await tail(fs, '/work/notes.md', 2)).lines).toEqual(['c', 'd']);
});

it('tail on a missing file returns an empty lines array', async () => {
  expect((await tail(fs, '/work/missing.md')).lines).toEqual([]);
});
```

(Reuse whatever fake-`DocFs`/fixture helper `coreutils.test.ts`'s existing `head` tests already use — read them first, do not invent a new one.)

- [ ] **Step 3: Run, verify fail.**

- [ ] **Step 4: Implement `tail` in `coreutils.ts`**

```ts
export async function tail(fs: DocFs, path: string, n = 10): Promise<{ lines: string[] }> {
  const v = await fs.readFile(path);
  if (!v) return { lines: [] };
  const allLines = v.text.split('\n');
  return { lines: allLines.slice(-n) };
}
```

(Match this to `head`'s EXACT null-handling/splitting behavior — read `head`'s real implementation first; the snippet above is illustrative, not necessarily byte-identical to `head`'s actual approach to trailing-newline handling. Verify with a test: `tail` and `head` should agree on line-splitting semantics for the same file.)

- [ ] **Step 5: Run, verify pass.**

- [ ] **Step 6: Wire the grammar layer (mirrors `find`'s Task 2 wiring from the prior plan exactly)**

`READ_VERBS`: add `'tail',`.
`ParsedCommand` union: add `| { verb: 'tail'; path: string; n?: number }`.
`ParsedCommandSchema`: add `z.object({ verb: z.literal('tail'), path: z.string(), n: z.number().optional() })` — **do this in the SAME commit as the grammar change**; this repo's `ls` work found a missing schema branch here is invisible to `tsc` (the schema is annotated `z.ZodType<ParsedCommand>`, checked only covariantly) and only surfaces via a runtime `.parse()` test. Add a regression test asserting `ParsedCommandSchema.parse(parseCommandLine('tail /work/notes.md'))` round-trips without throwing, and add `tail` to the existing "round-trips every command shape" test array if one exists.
`parseCommandLine` switch:
```ts
    case 'tail': {
      if (rest === '') return { error: 'tail needs a path — usage: tail <path> [n]' };
      const [path, nStr] = rest.split(/\s+/, 2);
      const n = nStr ? Number(nStr) : undefined;
      if (nStr !== undefined && (n === undefined || Number.isNaN(n))) {
        return { error: 'tail: n must be a number' };
      }
      return { verb: 'tail', path: path!, ...(n !== undefined ? { n } : {}) };
    }
```
`grammarFor`/advertisement entry: mirror whatever `find`'s entry looks like there.
`capability.ts`'s `ReadVerbSchema`: check whether it needs `'tail'` added (it needed `'ls'` and `'find'` both times in the prior plan — check first per this plan's Global Constraints rather than assuming, but budget for it being required).

- [ ] **Step 7: Run, verify fail then pass at the grammar layer** (same fail→implement→pass rhythm as Steps 2-5, applied to `command-grammar.test.ts`).

- [ ] **Step 8: Wire the compiler layer (`command-protocol.ts`)**

`ReadIntent` union: add `| { read: 'tail'; path: string; n?: number }`.
`compileCommand` switch:
```ts
    case 'tail':
      return { kind: 'read', intent: { read: 'tail', path: cmd.path, ...(cmd.n !== undefined ? { n: cmd.n } : {}) } };
```

- [ ] **Step 9: Wire execution (`assist-session.ts`)**

Add `'tail'` to `READ_COMMAND_VERBS` (the THIRD hand-maintained verb set this repo's Task 2 found — both `ls` and `find` needed adding here or they're unreachable through the real `runCommands()` loop despite passing unit tests on `runReadIntent` directly; write a full-loop regression test for `tail`, not just a `runReadIntent`-level test, mirroring whatever full-loop test the prior plan's Task 2 added for `ls`/`find`).

```ts
        case 'tail': {
          try {
            const { lines } = await docFsTail(this.docFs, intent.path, intent.n);
            return { label: `tail ${intent.path}`, result: lines.map((text) => ({ text })) };
          } catch (err) {
            return { label: `tail ${intent.path}`, result: { error: errMsg(err) } };
          }
        }
```

Import `tail as docFsTail` from `./docfs/index.js` alongside the existing `docFsLs`/`docFsFind` imports.

- [ ] **Step 10: Command help + regenerate the skill catalog**

Read `packages/contracts/src/command-help.ts`'s `ls`/`find` entries (added in the prior plan) for the exact tone/format, add a `tail` entry via `genericRead(...)` (check its real signature first, per this repo's own convention).

```bash
timeout 60 bun run emit:language
git diff --stat skill/m365-surface-commander/scripts/m365-cli-1.0.json skill/m365-surface-commander/references/generated-*.md
```
Read the actual diff (not just `--stat`) and confirm the ONLY change is the `tail` addition. If anything else changed, STOP and investigate before proceeding — an unexpected diff here means something upstream in this task altered the manifest unintentionally (this repo's prior plan found and fixed two real "unbuildable from git alone" bugs this exact check surfaced — do not skip it).

- [ ] **Step 11: Python skill-side parity**

Read `skill/m365-surface-commander/scripts/parse_commands.py`'s `parse_line` function and mirror the existing one-optional-arg pattern (`find` is the closest precedent: one required path, one optional trailing token) for `tail`. Match the TS `parseCommandLine`'s error text verbatim (this repo's prior plan requires byte-identical error strings between TS and Python — verify by literally comparing both strings side by side, not by memory). Add `tail` cases to whichever file Step 1's investigation in the PRIOR plan found actually owns commander-verb parity (`parity_corpus_test.py`, not `parity_test.py` — confirmed in the prior plan's Task 4) — add corpus entries to `golden-corpus.jsonl` and a `tail` mention to `capability-map.md`, following that exact precedent.

```bash
timeout 60 python3 skill/m365-surface-commander/scripts/parse_commands.py --self-test
timeout 60 python3 skill/parity_test.py
timeout 60 python3 skill/parity_corpus_test.py
timeout 90 python3 skill/validate_skill_bundles.py --check-zip
```

- [ ] **Step 12: Full gate + commit**

```bash
timeout 150 bun run typecheck
timeout 90 npx vitest run packages/contracts/src/command-grammar.test.ts packages/runtime/src/command-protocol.test.ts packages/runtime/src/assist-session.test.ts packages/runtime/src/docfs/coreutils.test.ts
timeout 30 npx prettier --check <every TS file touched above>
git add <every file touched above>
git commit -m "feat(runtime): add tail read verb (DocFs file-level, distinct from the pipeline tail transform)"
```

---

### Task 4: `cp`/`mv`/`rm` — `/work` artifact lifecycle

**Files:**
- Modify: `packages/runtime/src/workspace.ts` (new `cp`/`mv`/`rm` methods on `WorkspaceStore`)
- Modify: `packages/runtime/src/command-protocol.ts` (`WorkspaceIntent` union, `compileCommand`)
- Modify: `packages/contracts/src/command-grammar.ts` (grammar for `cp <src> <dst>`, `mv <src> <dst>`, `rm <name>`)
- Modify: `packages/contracts/src/command-help.ts`
- Test: `packages/runtime/src/workspace.test.ts`, `packages/contracts/src/command-grammar.test.ts`, `packages/runtime/src/command-protocol.test.ts`

**Interfaces:**
- Consumes: `WorkspaceStore`'s existing `artifacts: Map<id, WorkspaceArtifact>` and `aliases: Map<name, id>` (read `save`'s implementation in `workspace.ts` first — Task 1's brief in this document already excerpts it).
- Produces: `WorkspaceStore.cp(src: string, dst: string): WorkspaceResult`, `.mv(src: string, dst: string): WorkspaceResult`, `.rm(name: string): WorkspaceResult`; three new `WorkspaceIntent` variants (`{workspace:'cp', src, dst}`, `{workspace:'mv', src, dst}`, `{workspace:'rm', name}`); three new `ParsedCommand`/grammar verbs `cp`, `mv`, `rm`.

**Global Constraint reminder:** these operate ONLY on `/work`. There is no `/doc` equivalent — do not add a generic DocFs-level cp/mv/rm; this is a `WorkspaceStore` capability exposed through the SAME side-channel `save`/`cat`/`grep` already use (`WorkspaceIntent`, not `ReadIntent`/`DocFs`).

- [ ] **Step 1: Read `WorkspaceStore.save()`/`.cat()`/`.grep()` and their `WorkspaceIntent`/`compileCommand` wiring in full**

Confirms exactly how a ref (`ws:N` id or name-alias) resolves to an artifact today, so `cp`/`mv`/`rm` reuse that resolution rather than reinventing it. Also check `evictOldest()` (referenced in `save()`) to understand this store's existing capacity-bounding behavior — `cp` must interact correctly with it (a copy counts as a new artifact for eviction purposes).

- [ ] **Step 2: Write the failing `WorkspaceStore` tests**

```ts
it('cp duplicates an artifact under a new name with a new id', () => {
  const store = new WorkspaceStore();
  store.save({ name: 'a.tsv', sourceLabel: 'test', content: 'x\ty\n1\t2\n', kind: 'tsv' });
  const result = store.cp('a.tsv', 'b.tsv');
  expect(result.workspace).toBe('cp');
  const b = store.get('b.tsv');
  expect(b?.text).toBe(store.get('a.tsv')!.text);
  expect(b?.id).not.toBe(store.get('a.tsv')!.id);
});

it('cp on a missing source returns a corrective error', () => {
  const store = new WorkspaceStore();
  expect(store.cp('nope', 'x')).toHaveProperty('error');
});

it('mv renames an artifact in place (same id, old name no longer resolves)', () => {
  const store = new WorkspaceStore();
  store.save({ name: 'a.tsv', sourceLabel: 'test', content: '1', kind: 'tsv' });
  const id = store.get('a.tsv')!.id;
  store.mv('a.tsv', 'b.tsv');
  expect(store.get('b.tsv')?.id).toBe(id);
  expect(store.get('a.tsv')).toBeUndefined();
});

it('rm deletes an artifact so it no longer resolves by name or id', () => {
  const store = new WorkspaceStore();
  store.save({ name: 'a.tsv', sourceLabel: 'test', content: '1', kind: 'tsv' });
  const id = store.get('a.tsv')!.id;
  store.rm('a.tsv');
  expect(store.get('a.tsv')).toBeUndefined();
  expect(store.get(id)).toBeUndefined();
});

it('rm on a missing artifact returns a corrective error', () => {
  const store = new WorkspaceStore();
  expect(store.rm('nope')).toHaveProperty('error');
});

it('mv/cp onto an existing destination name overwrites it (last-write-wins, matching save() semantics)', () => {
  const store = new WorkspaceStore();
  store.save({ name: 'a.tsv', sourceLabel: 'test', content: '1', kind: 'tsv' });
  store.save({ name: 'b.tsv', sourceLabel: 'test', content: '2', kind: 'tsv' });
  store.cp('a.tsv', 'b.tsv');
  expect(store.get('b.tsv')?.text).toBe('1');
});
```

(Verify the last test's assumption — "overwrites, last-write-wins" — against `save()`'s actual `this.aliases.set(input.name, id)` behavior in Step 1 before asserting it; if `save()` throws/errors on a name collision instead, match THAT behavior instead of what's written above.)

- [ ] **Step 3: Run, verify fail.**

- [ ] **Step 4: Implement `cp`/`mv`/`rm` on `WorkspaceStore`**

Sketch (adapt field/type names exactly to what Step 1 found — `WorkspaceArtifact`'s real shape, `WorkspaceResult`'s real discriminated-union shape):

```ts
cp(src: string, dst: string): WorkspaceResult {
  const source = this.get(src);
  if (!source) return { error: `no such artifact: ${src}` };
  const id = `ws:${this.nextId++}`;
  const copy: WorkspaceArtifact = { ...source, id, name: dst, createdAt: new Date().toISOString() };
  this.artifacts.set(id, copy);
  this.aliases.set(dst, id);
  this.evictOldest();
  return { workspace: 'cp', artifact: summaryOf(copy) };
}

mv(src: string, dst: string): WorkspaceResult {
  const source = this.get(src);
  if (!source) return { error: `no such artifact: ${src}` };
  source.name = dst;
  this.aliases.delete(src);
  this.aliases.set(dst, source.id);
  return { workspace: 'mv', artifact: summaryOf(source) };
}

rm(name: string): WorkspaceResult {
  const artifact = this.get(name);
  if (!artifact) return { error: `no such artifact: ${name}` };
  this.artifacts.delete(artifact.id);
  this.aliases.delete(name);
  return { workspace: 'rm', name };
}
```

(These are sketches — `get()` may resolve by id OR name; `mv`'s `this.aliases.delete(src)` is only correct if `src` was itself a name and not an id someone renamed via its id — read `get()`'s actual resolution order in Step 1 and adjust. Also confirm `WorkspaceResult`'s discriminated union in `command-protocol.ts`/wherever it's defined needs new `'cp'`/`'mv'`/`'rm'` members before these compile.)

- [ ] **Step 5: Run, verify pass.**

- [ ] **Step 6: Wire `WorkspaceIntent` + `compileCommand`** (mirror `save`'s exact wiring pattern in `command-protocol.ts`):

```ts
  | { workspace: 'cp'; src: string; dst: string }
  | { workspace: 'mv'; src: string; dst: string }
  | { workspace: 'rm'; name: string };
```

- [ ] **Step 7: Wire the grammar** (`command-grammar.ts`) for three new verbs — `cp <src> <dst>`, `mv <src> <dst>`, `rm <name>` — each rejecting a missing required argument the same way `save`/`cat` already do (read their exact error-message style and match it). Add to `WORKSPACE_VERBS` (not `READ_VERBS` — these are workspace-mutation verbs, matching `save`'s classification, not read verbs). Add `ParsedCommandSchema` branches for all three in the SAME commit (per this plan's Global Constraints on the `ls` schema-branch gotcha).

- [ ] **Step 8: Command help + regenerate the skill catalog** (same procedure as Task 3 Step 10 — read the actual diff, confirm it's exactly the 3 new entries plus nothing else).

- [ ] **Step 9: Python skill-side parity** (same procedure as Task 3 Step 11, for `cp`/`mv`/`rm`).

- [ ] **Step 10: Full gate + commit**

```bash
timeout 150 bun run typecheck
timeout 90 npx vitest run packages/runtime/src/workspace.test.ts packages/contracts/src/command-grammar.test.ts packages/runtime/src/command-protocol.test.ts
timeout 30 npx prettier --check <every file touched>
git add <every file touched>
git commit -m "feat(runtime): add cp/mv/rm workspace artifact lifecycle verbs"
```

---

### Task 5: `/work` pseudo-directory rendering (naming-convention namespacing, no `mkdir`)

**Files:**
- Modify: `packages/runtime/src/docfs/work-mount.ts` (`readdir`, `find`'s glob matching path if it iterates artifact names directly rather than delegating to a shared `readdir`)
- Test: `packages/runtime/src/docfs/work-mount.test.ts`

**Interfaces:**
- Consumes: `WorkspaceStore.list()` (existing).
- Produces: `workMount(store).readdir(rel)` groups artifact names containing `/` under `rel` into `{kind:'dir'}` entries for the FIRST path segment past `rel`, and returns only the artifacts whose name is directly under `rel` as `{kind:'file'}` entries — exactly like `ls` on a real filesystem, but backed by nothing more than string-splitting `WorkspaceStore`'s existing flat name list. No new persisted state.

- [ ] **Step 1: Read `work-mount.ts`'s current `readdir` implementation in full**, and `docMount`'s `readdir` (the `/doc` mount) for a REFERENCE pattern of how directory-vs-file entries are already distinguished there (it already returns both `{kind:'dir'}` mount-prefix entries and `{kind:'file'}` entries at the DocFs root — reuse that same `DirEntry` shape, not a new one).

- [ ] **Step 2: Write the failing tests**

```ts
it('readdir("") groups names containing "/" into one dir entry per top-level prefix', async () => {
  const store = new WorkspaceStore();
  store.save({ name: 'q3-review/sales.tsv', sourceLabel: 'test', content: 'a', kind: 'tsv' });
  store.save({ name: 'q3-review/summary.md', sourceLabel: 'test', content: 'b', kind: 'md' });
  store.save({ name: 'flat.tsv', sourceLabel: 'test', content: 'c', kind: 'tsv' });
  const entries = await workMount(store).readdir('');
  expect(entries).toContainEqual({ name: 'q3-review', kind: 'dir' });
  expect(entries).toContainEqual(expect.objectContaining({ name: 'flat.tsv', kind: 'file' }));
  expect(entries.find((e) => e.name === 'q3-review/sales.tsv')).toBeUndefined();
});

it('readdir("q3-review") lists only the artifacts directly under that prefix, names relative to it', async () => {
  const store = new WorkspaceStore();
  store.save({ name: 'q3-review/sales.tsv', sourceLabel: 'test', content: 'a', kind: 'tsv' });
  const entries = await workMount(store).readdir('q3-review');
  expect(entries).toContainEqual(expect.objectContaining({ name: 'sales.tsv', kind: 'file' }));
});

it('readFile still resolves a namespaced artifact by its full name', async () => {
  const store = new WorkspaceStore();
  store.save({ name: 'q3-review/sales.tsv', sourceLabel: 'test', content: 'hello', kind: 'tsv' });
  const v = await workMount(store).readFile('q3-review/sales.tsv');
  expect(v?.text).toBe('hello');
});
```

- [ ] **Step 3: Run, verify fail.**

- [ ] **Step 4: Implement.** In `readdir(rel)`: partition `store.list()`'s names by whether they start with `rel === '' ? '' : rel + '/'`; for names with a further `/` beyond that prefix, collapse to one `{kind:'dir'}` entry per distinct next segment (dedup); for names with no further `/`, emit a `{kind:'file'}` entry with the name relative to `rel`. `readFile`/`stat`/`search` need NO changes if they already resolve by the artifact's full stored name (verify this in Step 1 — if `readFile` currently assumes flat names with some other lookup mechanism, adjust minimally to keep working for both flat and namespaced names).

- [ ] **Step 5: Run, verify pass.**

- [ ] **Step 6: Full gate + commit**

```bash
timeout 90 bun run typecheck
timeout 60 npx vitest run packages/runtime/src/docfs/work-mount.test.ts packages/runtime/src/docfs/coreutils.test.ts
timeout 30 npx prettier --check packages/runtime/src/docfs/work-mount.ts packages/runtime/src/docfs/work-mount.test.ts
git add packages/runtime/src/docfs/work-mount.ts packages/runtime/src/docfs/work-mount.test.ts
git commit -m "feat(runtime): render '/' in workspace artifact names as pseudo-directories in ls/find"
```

---

## Roadmap note (out of scope for this plan, tracked separately)

- Standalone `glob` verb and multi-artifact batch `cat`/`grep` — deferred with the existing cat/grep-over-DocFs unification (repo owner's explicit choice this session), pending the incident investigation's capture-legibility findings.
- A general expression language for computed columns beyond `derive`'s single binary arithmetic op (chained expressions, functions) — not requested; `derive` intentionally stays minimal.
- `rmdir`/empty-directory semantics — moot, since Task 5 introduces no persisted directory state to begin with.

## Self-Review

- **Spec coverage:** sed (Task 1), derive/awk-equivalent (Task 2), tail (Task 3), cp/mv/rm (Task 4), namespacing (Task 5) — all five items from the "bring these capabilities" request are covered; `glob`/`mkdir` are explicitly and deliberately not (per the repo owner's own three decisions this plan opens with).
- **Placeholder scan:** all steps carry concrete test code and implementation sketches. Where an implementation sketch is explicitly marked "adapt to the real X" (Task 4 Step 4, Task 3 Step 4), that's deliberate scene-setting pointing at a real, findable existing API whose exact shape the plan author hasn't independently re-verified line-by-line — not an unresolved requirement; the tests in the same task pin the required OBSERVABLE behavior precisely regardless of the exact internal field names used to get there.
- **Type consistency:** `WorkspaceIntent`'s new `cp`/`mv`/`rm` variants use `src`/`dst`/`name` field names consistent with the grammar verbs' own argument names (Task 4). `ReadIntent`'s `tail` variant reuses `path`/`n` field names consistent with `find`'s `path`/`glob` precedent (string path, optional second field) from the prior plan.
