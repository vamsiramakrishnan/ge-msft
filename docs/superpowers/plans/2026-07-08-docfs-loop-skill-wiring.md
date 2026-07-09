# Conductor Phase 2a: DocFs Verbs, Skill Parity, Discovery Wiring, No-Fence Telemetry — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give the model two new DocFs-backed read verbs (`ls`, `find`) reachable end-to-end — grammar → compiler → execution → the auto-generated skill catalog → the Python skill's own parser — plus make the skill/data-store discovery already built (`compose.ts`'s `availableAgents`/`availableDataStores`) actually reach the UI layer instead of being computed and only logged, plus add raw-text diagnostics on `no-fence` events so a repeat of the live incident (a commander turn whose reply doesn't parse as a `cmd` fence) is diagnosable from telemetry instead of transcript archaeology.

**Architecture:** `ls`/`find` are NEW, purely additive `ReadIntent` variants that execute against a `DocFs` instance (`createDocFs({bridge, workspace})`, built in Phase 1) already spanning `/doc` and `/work` — they do not touch or refactor the existing `outline`/`read`/`search`/`workspace`/`cat`/`grep` verbs, which keeps this phase's regression risk to zero on the live incident's code path. The skill-side catalog (`generated-capability-catalog.md`, `generated-command-catalog.md`, `m365-cli-1.0.json`) is auto-generated from `packages/contracts/src/language-manifest.ts` via `bun run emit:language` — adding the two verbs there and re-running the emitter is what "wires it into the skill" for documentation purposes; `skill/m365-surface-commander/scripts/parse_commands.py` is a SEPARATE, hand-maintained mirror the skill runs server-side to self-validate its own draft output, so it needs its own matching update. Discovery wiring threads `ComposedSession.availableAgents`/`availableDataStores` (already computed, currently dead-ended at a console log) into `PanelController`'s public state, the natural seam a future `@`-picker renders from.

**Tech Stack:** TypeScript (Bun workspaces, Vitest, Zod), Python 3 stdlib (skill-side parser/parity, no new deps).

## Global Constraints

- `bun run typecheck` clean · `npx vitest run <touched test files>` green · `npx prettier --check <touched files>` clean, every task.
- Additive only: do not modify the existing behavior of `outline`, `read`, `search`, `workspace`, `save`, `cat`, `grep` verbs — those stay exactly as they are today. `ls`/`find` are new, separate `ReadIntent` variants.
- `ls`/`find` are READ-ONLY (no write/mutation) — they only ever call `DocFs.readdir`/coreutils `ls`/`find`, never `DocBridge.actuate`.
- Regenerating the language manifest (`bun run emit:language`) must produce the exact files listed in Task 3 with NO unrelated diff — if it changes anything besides the `ls`/`find` additions, stop and investigate before committing (per `emit-language-manifest.mjs`'s own comment: "a byte change is a real language change").
- Python changes are stdlib-only (no new pip deps), matching the existing skill scripts.
- This plan continues on branch `feat/conductor-docfs` (Phase 1 of the conductor roadmap already lives there, reviewed "Merge with follow-ups"). Do not create a new branch.

---

## File Structure

- Modify `packages/contracts/src/command-grammar.ts` — `ParsedCommand` union, `parseCommandLine`, `READ_VERBS`.
- Modify `packages/runtime/src/command-protocol.ts` — `ReadIntent` union, `compileCommand`.
- Modify `packages/runtime/src/assist-session.ts` — add a `docFs` field; `runReadIntent`'s `ls`/`find` cases; no-fence telemetry.
- Modify `packages/contracts/src/command-help.ts` — `COMMAND_HELP` entries for `ls`/`find`.
- Regenerate (via script, not by hand) `skill/m365-surface-commander/scripts/m365-cli-1.0.json`, `skill/m365-surface-commander/references/generated-capability-catalog.md`, `skill/m365-surface-commander/references/generated-command-catalog.md`.
- Modify `skill/m365-surface-commander/scripts/parse_commands.py` — recognize `ls`/`find` syntax.
- Modify `skill/parity_test.py` (or add a sibling check) — pin the new verbs into the lockstep guard.
- Modify `packages/web-shell/src/controller.ts` — expose `availableAgents`/`availableDataStores` in `PanelState`.
- Modify `packages/web-shell/src/taskpane/main.tsx` — pass discovery results into the controller instead of only logging counts.

---

## Task 1: `ls` read verb — grammar, compiler, execution

**Files:**
- Modify: `packages/contracts/src/command-grammar.ts` (`ParsedCommand` union ~line 146-176, `parseCommandLine` ~line 400, `READ_VERBS` ~line 50)
- Modify: `packages/runtime/src/command-protocol.ts` (`ReadIntent` union ~line 34-44, `compileCommand` ~line 85-92)
- Modify: `packages/runtime/src/assist-session.ts` (add `docFs` field near line 396-404, `runReadIntent` switch ~line 1533-1536)
- Test: `packages/contracts/src/command-grammar.test.ts`, `packages/runtime/src/command-protocol.test.ts`, `packages/runtime/src/assist-session.test.ts`

**Interfaces:**
- Consumes: `createDocFs`, `ls` from `@ge/runtime`'s `docfs/index.ts` (Phase 1, already merged on this branch — `createDocFs({bridge, workspace}): DocFs`; `ls(fs: DocFs, path: string): Promise<string[]>`).
- Produces: `ParsedCommand` variant `{ verb: 'ls'; path: string }`; `ReadIntent` variant `{ read: 'ls'; path: string }`; `AssistSession` gains a `private readonly docFs: DocFs` field other tasks (Task 2) also use.

- [ ] **Step 1: Write the failing grammar test**

Add to `packages/contracts/src/command-grammar.test.ts` (find the `describe('parseCommandLine'` block and add nearby the existing `outline`/`read` cases):

```ts
it('parses ls with a path', () => {
  expect(parseCommandLine('ls /doc')).toEqual({ verb: 'ls', path: '/doc' });
});

it('rejects ls with no path', () => {
  const result = parseCommandLine('ls');
  expect(result).toHaveProperty('error');
});
```

- [ ] **Step 2: Run it, verify it fails**

Run: `npx vitest run packages/contracts/src/command-grammar.test.ts -t "parses ls"`
Expected: FAIL — `parseCommandLine('ls /doc')` currently returns the "did-you-mean unknown verb" error path, not `{ verb: 'ls', path: '/doc' }`.

- [ ] **Step 3: Add the grammar support**

In `packages/contracts/src/command-grammar.ts`, add to the `ParsedCommand` union (next to the other read verbs, e.g. right after `| { verb: 'read'; selector: string }`):

```ts
  | { verb: 'ls'; path: string }
```

In `parseCommandLine`'s verb switch, add a case near `case 'read':`:

```ts
    case 'ls': {
      if (rest === '') return { error: 'ls needs a path — usage: ls <path>, e.g. ls /doc' };
      return { verb: 'ls', path: rest };
    }
```

Add `'ls'` to the `READ_VERBS` array (~line 50), in the same list as `'outline'`/`'read'`/`'search'`:

```ts
export const READ_VERBS = [
  'outline',
  'read',
  'search',
  'ls',
  // ... existing entries continue unchanged
```

- [ ] **Step 4: Run the grammar test, verify it passes**

Run: `npx vitest run packages/contracts/src/command-grammar.test.ts -t "ls"`
Expected: PASS (both the parse and the missing-path-error case).

- [ ] **Step 5: Write the failing compiler test**

Add to `packages/runtime/src/command-protocol.test.ts` (find `describe('compileCommand'`):

```ts
it('compiles ls to a read intent', () => {
  const ctx = { surface: 'excel' as const, mintChangeId: () => asChangeId('c1') };
  const compiled = compileCommand({ verb: 'ls', path: '/doc' }, ctx);
  expect(compiled).toEqual({ kind: 'read', intent: { read: 'ls', path: '/doc' } });
});
```

(Match whatever helper the file already uses to construct a `ctx`/`ChangeId` for other `compileCommand` tests — copy an existing test's setup verbatim rather than inventing new helpers.)

- [ ] **Step 6: Run it, verify it fails**

Run: `npx vitest run packages/runtime/src/command-protocol.test.ts -t "compiles ls"`
Expected: FAIL — TypeScript will actually fail to compile until the `ReadIntent`/`compileCommand` case exists, so this step's failure is a type error, not a runtime assertion failure; that is expected.

- [ ] **Step 7: Add the compiler support**

In `packages/runtime/src/command-protocol.ts`, add to the `ReadIntent` union (next to `| { read: 'range'; selector: string }`):

```ts
  | { read: 'ls'; path: string } // → DocFs.readdir() via ls() coreutil
```

In `compileCommand`'s switch, add a case near `case 'read':`:

```ts
    case 'ls':
      return { kind: 'read', intent: { read: 'ls', path: cmd.path } };
```

- [ ] **Step 8: Run the compiler test, verify it passes**

Run: `npx vitest run packages/runtime/src/command-protocol.test.ts -t "compiles ls"`
Expected: PASS.

- [ ] **Step 9: Write the failing execution test**

Add to `packages/runtime/src/assist-session.test.ts`. First find how an existing test constructs an `AssistSession` with a fake bridge (search the file for `new AssistSession(`) and copy that setup. Then add:

```ts
it('runs the ls read intent against DocFs (/work)', async () => {
  const session = new AssistSession(fakeBridge(), fakeClient(), {}); // match the exact fake-bridge/fake-client helpers already used elsewhere in this file
  const { result } = await (session as any).runReadIntent({ read: 'ls', path: '/work' });
  // /work is empty (no artifacts saved yet in this session) — ls returns an empty listing, not an error.
  expect(result).not.toHaveProperty('error');
});
```

(If the file exposes `runReadIntent` differently — e.g. only reachable through `runCommands`/`ask` — replace this with an integration-style test that runs `ls /work` as a full command turn instead of reaching into the private method. Check the file's existing patterns for testing other read verbs like `search`/`outline` and mirror that pattern exactly; do not invent a new testing approach for this one verb.)

- [ ] **Step 10: Run it, verify it fails**

Run: `npx vitest run packages/runtime/src/assist-session.test.ts -t "ls read intent"`
Expected: FAIL — `runReadIntent`'s switch has no `'ls'` case yet, so it falls through to the switch's default/exhaustiveness handling (a TypeScript error or a runtime "unhandled read intent" branch, depending on how the switch is written — check `runReadIntent`'s existing structure for its exhaustiveness pattern before writing this step's expected failure mode precisely).

- [ ] **Step 11: Wire the docFs field and the execution case**

In `packages/runtime/src/assist-session.ts`, add the import (near the existing imports from `./workspace.js`):

```ts
import { createDocFs, ls as docFsLs, find as docFsFind, type DocFs } from './docfs/index.js';
```

Add a field declaration (next to `private readonly workspace = new WorkspaceStore();` at line 362) — declare WITHOUT an initializer to avoid any class-field/constructor-parameter-property ordering ambiguity:

```ts
  private readonly docFs: DocFs;
```

In the constructor body (after the existing three assignment lines, i.e. after `this.compaction = { ...DEFAULT_COMPACTION, ...options.compaction };`), add:

```ts
    this.docFs = createDocFs({ bridge, workspace: this.workspace });
```

(Uses the constructor's own `bridge` parameter directly — not `this.bridge` — and `this.workspace`, which is guaranteed initialized by the time the constructor body runs since class field initializers run before the constructor body executes.)

In `runReadIntent`'s switch (the `case 'outline':` / `case 'range':` / `case 'search':` block, ~line 1536 onward), add:

```ts
        case 'ls': {
          try {
            const lines = await docFsLs(this.docFs, intent.path);
            return { label: `ls ${intent.path}`, result: lines.map((text) => ({ text })) };
          } catch (err) {
            return { label: `ls ${intent.path}`, result: { error: errMsg(err) } };
          }
        }
```

(`errMsg` is already used elsewhere in this file for error formatting in `runWorkspaceIntent`'s catch block — reuse it, don't redefine it. The result shape `lines.map((text) => ({ text }))` matches the existing array-of-`{text}` convention `readResultToText` already handles — no changes needed to `readResultToText`/`isReadErrorResult`.)

- [ ] **Step 12: Run the execution test, verify it passes**

Run: `npx vitest run packages/runtime/src/assist-session.test.ts -t "ls read intent"`
Expected: PASS.

- [ ] **Step 13: Full gate + commit**

```bash
bun run typecheck
npx vitest run packages/contracts/src/command-grammar.test.ts packages/runtime/src/command-protocol.test.ts packages/runtime/src/assist-session.test.ts
npx prettier --check packages/contracts/src/command-grammar.ts packages/contracts/src/command-grammar.test.ts packages/runtime/src/command-protocol.ts packages/runtime/src/command-protocol.test.ts packages/runtime/src/assist-session.ts packages/runtime/src/assist-session.test.ts
git add packages/contracts/src/command-grammar.ts packages/contracts/src/command-grammar.test.ts packages/runtime/src/command-protocol.ts packages/runtime/src/command-protocol.test.ts packages/runtime/src/assist-session.ts packages/runtime/src/assist-session.test.ts
git commit -m "feat(runtime): add ls read verb, DocFs-backed"
```

---

## Task 2: `find` read verb — grammar, compiler, execution

**Files:** same files as Task 1, plus their test files.

**Interfaces:**
- Consumes: `find(fs: DocFs, path: string, glob?: string): Promise<string[]>` from `./docfs/index.js` (already imported in Task 1's step 11 as `docFsFind`); `this.docFs` (Task 1).
- Produces: `ParsedCommand` variant `{ verb: 'find'; path: string; glob?: string }`; `ReadIntent` variant `{ read: 'find'; path: string; glob?: string }`.

- [ ] **Step 1: Write the failing grammar tests**

Add to `packages/contracts/src/command-grammar.test.ts`:

```ts
it('parses find with a path only', () => {
  expect(parseCommandLine('find /work')).toEqual({ verb: 'find', path: '/work' });
});

it('parses find with a path and a glob', () => {
  expect(parseCommandLine('find /work *.tsv')).toEqual({
    verb: 'find',
    path: '/work',
    glob: '*.tsv',
  });
});

it('rejects find with no path', () => {
  const result = parseCommandLine('find');
  expect(result).toHaveProperty('error');
});
```

- [ ] **Step 2: Run, verify fail**

Run: `npx vitest run packages/contracts/src/command-grammar.test.ts -t "find"`
Expected: FAIL (unknown verb).

- [ ] **Step 3: Add grammar support**

`ParsedCommand` union, next to the `ls` entry from Task 1:

```ts
  | { verb: 'find'; path: string; glob?: string }
```

`parseCommandLine` switch, next to `case 'ls':`:

```ts
    case 'find': {
      if (rest === '') return { error: 'find needs a path — usage: find <path> [glob]' };
      const [path, glob] = rest.split(/\s+/, 2);
      return { verb: 'find', path: path!, ...(glob ? { glob } : {}) };
    }
```

`READ_VERBS`, next to `'ls'`:

```ts
  'find',
```

- [ ] **Step 4: Run, verify pass**

Run: `npx vitest run packages/contracts/src/command-grammar.test.ts -t "find"`
Expected: PASS.

- [ ] **Step 5: Write the failing compiler test**

Add to `packages/runtime/src/command-protocol.test.ts`:

```ts
it('compiles find to a read intent', () => {
  const ctx = { surface: 'excel' as const, mintChangeId: () => asChangeId('c1') };
  expect(compileCommand({ verb: 'find', path: '/work' }, ctx)).toEqual({
    kind: 'read',
    intent: { read: 'find', path: '/work' },
  });
  expect(compileCommand({ verb: 'find', path: '/work', glob: '*.tsv' }, ctx)).toEqual({
    kind: 'read',
    intent: { read: 'find', path: '/work', glob: '*.tsv' },
  });
});
```

- [ ] **Step 6: Run, verify fail** (type error until the union/case exist, same as Task 1 Step 6).

- [ ] **Step 7: Add compiler support**

`ReadIntent` union, next to `ls`:

```ts
  | { read: 'find'; path: string; glob?: string } // → DocFs coreutil find()
```

`compileCommand` switch, next to `case 'ls':`:

```ts
    case 'find':
      return { kind: 'read', intent: { read: 'find', path: cmd.path, ...(cmd.glob ? { glob: cmd.glob } : {}) } };
```

- [ ] **Step 8: Run, verify pass.**

- [ ] **Step 9: Write the failing execution test**

Add to `packages/runtime/src/assist-session.test.ts`, following the exact pattern Task 1 Step 9 established for `ls` (same fake-bridge/session setup):

```ts
it('runs the find read intent against DocFs (/work)', async () => {
  const session = new AssistSession(fakeBridge(), fakeClient(), {});
  const { result } = await (session as any).runReadIntent({ read: 'find', path: '/work' });
  expect(result).not.toHaveProperty('error');
});
```

- [ ] **Step 10: Run, verify fail** (no `'find'` case yet).

- [ ] **Step 11: Add the execution case**

In `runReadIntent`'s switch, next to the `case 'ls':` from Task 1:

```ts
        case 'find': {
          try {
            const paths = await docFsFind(this.docFs, intent.path, intent.glob);
            return { label: `find ${intent.path}`, result: paths.map((text) => ({ text })) };
          } catch (err) {
            return { label: `find ${intent.path}`, result: { error: errMsg(err) } };
          }
        }
```

- [ ] **Step 12: Run, verify pass.**

- [ ] **Step 13: Full gate + commit**

```bash
bun run typecheck
npx vitest run packages/contracts/src/command-grammar.test.ts packages/runtime/src/command-protocol.test.ts packages/runtime/src/assist-session.test.ts
npx prettier --check packages/contracts/src/command-grammar.ts packages/contracts/src/command-grammar.test.ts packages/runtime/src/command-protocol.ts packages/runtime/src/command-protocol.test.ts packages/runtime/src/assist-session.ts packages/runtime/src/assist-session.test.ts
git add packages/contracts/src/command-grammar.ts packages/contracts/src/command-grammar.test.ts packages/runtime/src/command-protocol.ts packages/runtime/src/command-protocol.test.ts packages/runtime/src/assist-session.ts packages/runtime/src/assist-session.test.ts
git commit -m "feat(runtime): add find read verb, DocFs-backed"
```

---

## Task 3: Command help entries + regenerate the skill catalog

**Files:**
- Modify: `packages/contracts/src/command-help.ts`
- Regenerate (do not hand-edit): `skill/m365-surface-commander/scripts/m365-cli-1.0.json`, `skill/m365-surface-commander/references/generated-capability-catalog.md`, `skill/m365-surface-commander/references/generated-command-catalog.md`
- Test: `packages/contracts/src/language-manifest.test.ts` (existing consistency assertions must still pass with the two new verbs present)

**Interfaces:**
- Consumes: `CommandHelpEntrySchema` shape (`command`, `useWhen`, `syntax`, `discovery`, `sequence`, `examples`, `doNot`, `failureModes`, `safety`) from `packages/contracts/src/command-help.ts:3-14`; `READ_VERBS` (Tasks 1-2, already includes `'ls'`/`'find'`).
- Produces: the regenerated skill-bundle JSON/docs later tasks (Task 5's parity test, and the skill's own runtime) read.

- [ ] **Step 1: Read the existing pattern**

Open `packages/contracts/src/command-help.ts` and read the `outline` entry (the `genericRead('outline', 'outline', ...)` call near line 17) plus its neighboring `search`/`read` entries in full, to match their exact tone/format (this file's entries are the literal text a model reads — get the style right, don't paraphrase).

- [ ] **Step 2: Add `ls` and `find` help entries**

In the same object literal `COMMAND_HELP` is built from, add two entries following the `genericRead(...)` helper's calling convention observed in Step 1 (adapt the exact call to that helper's real parameter list — do not guess its signature without reading it first in Step 1):

```ts
  ls: genericRead(
    'ls',
    'ls <path>',
    'you need to see what exists under /doc (the live document) or /work (saved artifacts) before reading one entry',
  ),
  find: genericRead(
    'find',
    'find <path> [glob]',
    'you need to locate an artifact or document entry by name pattern instead of listing everything',
  ),
```

(If `genericRead` takes additional required arguments beyond these three — check its definition in Step 1 — supply them consistent with how `outline`/`read`/`search` call it.)

- [ ] **Step 3: Run the manifest consistency test**

Run: `npx vitest run packages/contracts/src/language-manifest.test.ts`
Expected: PASS — `assertManifestConsistent` (called by `buildLanguageManifest`) cross-checks that every `READ_VERBS` entry has a `COMMAND_HELP` entry; this is the test that would fail if Step 2 were skipped.

- [ ] **Step 4: Regenerate the skill catalog**

```bash
bun run emit:language
```

Expected output includes lines like `wrote .../m365-cli-1.0.json (…, N write verbs)` and `wrote .../generated-capability-catalog.md`, `wrote .../generated-command-catalog.md`.

- [ ] **Step 5: Verify the regeneration is exactly the expected diff**

```bash
git diff --stat skill/m365-surface-commander/scripts/m365-cli-1.0.json skill/m365-surface-commander/references/generated-capability-catalog.md skill/m365-surface-commander/references/generated-command-catalog.md
```

Read the actual diff (`git diff` without `--stat`) and confirm the ONLY changes are the addition of `ls`/`find` (their JSON entries, and their rendered catalog sections). If anything else changed, STOP — per this plan's Global Constraints, an unexpected diff here means something in Tasks 1-2 altered the manifest in an unintended way; investigate before proceeding.

- [ ] **Step 6: Full gate + commit**

```bash
bun run typecheck
npx vitest run packages/contracts/src/language-manifest.test.ts
npx prettier --check packages/contracts/src/command-help.ts
git add packages/contracts/src/command-help.ts skill/m365-surface-commander/scripts/m365-cli-1.0.json skill/m365-surface-commander/references/generated-capability-catalog.md skill/m365-surface-commander/references/generated-command-catalog.md
git commit -m "feat(contracts): document ls/find; regenerate the skill catalog"
```

---

## Task 4: Python skill-side parity for `ls`/`find`

**Files:**
- Modify: `skill/m365-surface-commander/scripts/parse_commands.py`
- Modify: `skill/parity_test.py` (extend, or confirm an existing broader check already covers this — see Step 1)
- Test: run the Python scripts directly (stdlib `unittest`/plain `assert`-based, matching this repo's existing skill test style)

**Interfaces:**
- Consumes: the verb list now includes `ls`/`find` (Tasks 1-3); the exact syntax `ls <path>` / `find <path> [glob]` established in Task 1-2's grammar.
- Produces: a `parse_commands.py` that recognizes these two verbs when the skill self-validates a draft `cmd` block server-side.

- [ ] **Step 1: Read the current verb recognition**

Read `skill/m365-surface-commander/scripts/parse_commands.py`'s `parse_line` function (search for `def parse_line`) in full, and identify exactly how an existing no-argument-required-but-one-arg verb like `read` or `search` is recognized (what data structure holds the verb list, is it a dict/set of verb→arg-arity, a big if/elif chain, etc.). Also check `skill/parity_test.py`'s scope (read the whole file — it currently pins the PLANNER's `INTENTS`/`CONTEXT_HINTS` only, per its own docstring "pin the GE skills' planner verb set to the authoritative intent model"). Confirm whether a SEPARATE test already pins the COMMANDER's verb set (search `skill/` for other `*parity*` or `*_test.py` files, e.g. `skill/test_tooling.py`) — if one exists, that is where `ls`/`find` need to be added instead of `parity_test.py`. Report which file actually owns commander-verb parity before proceeding to Step 2.

- [ ] **Step 2: Add recognition for `ls`/`find` in `parse_commands.py`**

Using the exact pattern identified in Step 1 (mirror whatever data structure/dispatch `read`/`search` already use — do not introduce a new dispatch style for just these two verbs), add `ls` (one required arg: path) and `find` (one required arg: path, one optional arg: glob) to the same structure. The concrete edit depends on Step 1's finding; at minimum it must make `parse_commands.py` accept `"ls /doc"`, `"find /work"`, and `"find /work *.tsv"` as valid lines, and reject bare `"ls"` / `"find"` (no path) the same way `"read"`-with-no-selector or the equivalent existing required-arg verb is rejected today (match that exact error style).

- [ ] **Step 3: Add/extend the parity test**

In whichever file Step 1 identified as owning commander-verb parity (or, if none exists, add a new minimal check to `skill/parity_test.py` following its existing style — a plain function returning a list of failure strings, matching the file's own pattern), add an assertion that `ls` and `find` are both present in the recognized verb set, e.g. (adapt names/structures to what Step 1 found):

```python
EXPECTED_READ_VERBS = {..., "ls", "find"}  # extend whatever the existing constant is called
```

- [ ] **Step 4: Run the parity check**

```bash
python3 skill/parity_test.py
```

(Or whichever script Step 1 identified.) Expected: exits 0, no failures printed.

- [ ] **Step 5: Run the broader skill validation**

```bash
python3 skill/validate_skill_bundles.py --check-zip
```

Expected: passes (confirms the skill bundle's own internal consistency isn't broken by the parser change).

- [ ] **Step 6: Commit**

```bash
git add skill/m365-surface-commander/scripts/parse_commands.py skill/parity_test.py
git commit -m "feat(skill): recognize ls/find in the commander's self-check parser"
```

(If Step 1 found a different file owns commander parity, `git add` that file instead of/in addition to `parity_test.py`.)

---

## Task 5: No-fence raw-text diagnostic telemetry

**Files:**
- Modify: `packages/runtime/src/assist-session.ts` (the `no-fence` event site, ~line 761-763, and the `no-fence` event type, ~line 177)
- Test: `packages/runtime/src/command-protocol.test.ts` or `packages/runtime/src/assist-session.branches.test.ts` (whichever already covers `no-fence` — search for `'no-fence'` in test files first)

**Interfaces:**
- Consumes: `turnText` (the fully-accumulated model reply for the turn, already in scope at the `no-fence` yield site).
- Produces: the `no-fence` event's payload gains an optional diagnostic field; nothing downstream that pattern-matches on `{ type: 'no-fence' }` breaks (additive field only).

- [ ] **Step 1: Read the current no-fence event shape and its one consumer**

Read `packages/runtime/src/assist-session.ts:177` (the `no-fence` member of the loop-event union) and its yield site (~line 761-763). Read `packages/web-shell/src/controller.ts`'s `reduceLoopEvent` handling of `case 'no-fence':` (grep for `'no-fence'` in that file) to confirm it only reads `.turn` today — adding a new optional field must not require changes there (it will simply be ignored), but confirm this by reading the actual destructuring/usage before proceeding.

- [ ] **Step 2: Write the failing test**

Find the existing test(s) asserting on a `no-fence` event (grep `'no-fence'` in `packages/runtime/src/assist-session.branches.test.ts` and `command-protocol.test.ts`; there is at least one, per the file already containing `'two consecutive no-fence replies end the loop'`). Add a new assertion to that same test (or a new adjacent one) — for example, if the existing test constructs a fake client whose reply is some known non-fenced string like `'not a fence'`:

```ts
it('a no-fence event carries a bounded, redacted snippet of the unparsed reply', async () => {
  // Reuse this file's existing fake-client/session setup that produces a no-fence turn.
  const events = /* however this file already collects the event stream for the no-fence test */;
  const noFence = events.find((e) => e.type === 'no-fence');
  expect(noFence?.rawSnippet).toBeDefined();
  expect(noFence!.rawSnippet!.length).toBeLessThanOrEqual(200);
});
```

(Match this file's ACTUAL existing helper names for constructing the event stream and fake client — read the neighboring `'two consecutive no-fence replies end the loop'` test in full first and copy its setup exactly; do not invent new helpers.)

- [ ] **Step 3: Run it, verify it fails**

Run: `npx vitest run <the file from Step 2> -t "rawSnippet"`
Expected: FAIL — `rawSnippet` does not exist on the event yet (`undefined`).

- [ ] **Step 4: Add the field and populate it**

In the `no-fence` event type (~line 177):

```ts
  | { type: 'no-fence'; turn: number; rawSnippet: string }
```

At the yield site (~line 761-763), replace:

```ts
      if (!found) {
        yield { type: 'no-fence', turn };
```

with:

```ts
      if (!found) {
        yield { type: 'no-fence', turn, rawSnippet: redactedSnippet(turnText) };
```

Add the helper near the other private helper functions in this file (e.g. next to `noFenceReprompt`):

```ts
/**
 * A bounded, best-effort-redacted preview of a turn's unparsed reply, attached to `no-fence`
 * events so a repeat of a parse-miss is diagnosable from telemetry instead of requiring a full
 * transcript capture. Bounded to 200 chars total (100 head + 100 tail, the two ends most likely to
 * show WHY extraction failed — a missing opening fence or a missing closing one) and strips
 * anything that looks like a quoted string (a crude guard against accidentally logging pasted
 * document content the model may have echoed back).
 */
function redactedSnippet(text: string): string {
  const cleaned = text.replace(/"[^"]*"/g, '"…"').replace(/'[^']*'/g, "'…'");
  if (cleaned.length <= 200) return cleaned;
  return `${cleaned.slice(0, 100)}…${cleaned.slice(-100)}`;
}
```

- [ ] **Step 5: Run the test, verify it passes**

Run: `npx vitest run <the file from Step 2> -t "rawSnippet"`
Expected: PASS.

- [ ] **Step 6: Run the full existing no-fence test suite to confirm no regression**

Run: `npx vitest run packages/runtime/src/assist-session.branches.test.ts packages/runtime/src/command-protocol.test.ts`
Expected: all pass, including the pre-existing `'two consecutive no-fence replies end the loop'` test (the new field is additive, so its existing assertions on `.turn`/event count must be unaffected).

- [ ] **Step 7: Full gate + commit**

```bash
bun run typecheck
npx prettier --check packages/runtime/src/assist-session.ts <the test file from Step 2>
git add packages/runtime/src/assist-session.ts <the test file from Step 2>
git commit -m "feat(runtime): attach a bounded, redacted raw-text snippet to no-fence events"
```

---

## Task 6: Wire discovered skills/data stores into `PanelController` state

**Files:**
- Modify: `packages/web-shell/src/controller.ts`
- Modify: `packages/web-shell/src/taskpane/main.tsx`
- Test: `packages/web-shell/src/controller.test.ts`

**Interfaces:**
- Consumes: `ComposedSession.availableAgents: AgentView[]` and `.availableDataStores: EngineDataStore[]` (already computed in `compose.ts`, currently only logged in `main.tsx`); `AgentView`/`EngineDataStore` types from `@ge/gemini-client`.
- Produces: `PanelController`'s public state gains `availableAgents: AgentView[]` and `availableDataStores: EngineDataStore[]` fields (empty arrays by default), settable via a new method `setDiscoveredCatalog(agents, dataStores)`, which a future `@`-picker UI reads from `PanelController.getState()`.

- [ ] **Step 1: Read the current `PanelController` state shape**

Read `packages/web-shell/src/controller.ts`'s `PanelState` interface definition (grep `interface PanelState`) and one existing simple state field (e.g. `skills`, already present per `registerSkills(e){this.set({skills:e})}` referenced earlier in this codebase) to copy its exact pattern for adding a new settable field.

- [ ] **Step 2: Write the failing test**

Add to `packages/web-shell/src/controller.test.ts` (find how an existing simple setter like `registerSkills`/`listSkills` is tested and copy that pattern):

```ts
it('setDiscoveredCatalog stores available agents and data stores in state', () => {
  const controller = /* however this file's other tests construct a PanelController */;
  const agents = [{ id: 'a1', displayName: 'Test Skill' }] as any;
  const dataStores = [{ id: 'ds1', resourceName: 'r1', displayName: 'SP Files', connector: 'SharePoint' }] as any;
  controller.setDiscoveredCatalog(agents, dataStores);
  expect(controller.getState().availableAgents).toEqual(agents);
  expect(controller.getState().availableDataStores).toEqual(dataStores);
});

it('availableAgents/availableDataStores default to empty arrays', () => {
  const controller = /* same construction as above */;
  expect(controller.getState().availableAgents).toEqual([]);
  expect(controller.getState().availableDataStores).toEqual([]);
});
```

(Replace the placeholder construction comment with this file's actual existing controller-construction helper — read at least two existing tests in this file first and copy their setup exactly.)

- [ ] **Step 3: Run it, verify it fails**

Run: `npx vitest run packages/web-shell/src/controller.test.ts -t "setDiscoveredCatalog"`
Expected: FAIL — `setDiscoveredCatalog` does not exist; `availableAgents`/`availableDataStores` are not on the state.

- [ ] **Step 4: Add the state fields and setter**

In `packages/web-shell/src/controller.ts`:

Add the import (near the top, alongside other `@ge/gemini-client` type imports if any already exist, or add a new import line):

```ts
import type { AgentView, EngineDataStore } from '@ge/gemini-client';
```

In the `PanelState` interface, add:

```ts
  availableAgents: AgentView[];
  availableDataStores: EngineDataStore[];
```

Find wherever the initial/default `PanelState` object is constructed (search for where `skills: []` or similar defaults are set) and add matching defaults:

```ts
  availableAgents: [],
  availableDataStores: [],
```

Add the setter method, next to `registerSkills` (per Task's Interfaces section, this codebase already has that method — place this one immediately after it for discoverability):

```ts
  setDiscoveredCatalog(agents: AgentView[], dataStores: EngineDataStore[]): void {
    this.set({ availableAgents: agents, availableDataStores: dataStores });
  }
```

- [ ] **Step 5: Run the test, verify it passes**

Run: `npx vitest run packages/web-shell/src/controller.test.ts -t "setDiscoveredCatalog\|default to empty"`
Expected: PASS.

- [ ] **Step 6: Wire it from boot**

In `packages/web-shell/src/taskpane/main.tsx`, find the existing block (added in an earlier phase of this work) that logs discovery counts:

```ts
    recordAuthDebug('catalog.discovered', {
      agents: availableAgents.length,
      dataStores: availableDataStores.length,
      connectors: [...new Set(availableDataStores.map((d) => d.connector))].join(','),
    });
```

Immediately after it (same `try`/scope), add:

```ts
    controller.setDiscoveredCatalog(availableAgents, availableDataStores);
```

(`controller` is the `PanelController` instance already constructed a few lines below in this function today — if `setDiscoveredCatalog` is called before `controller` exists in the current code order, move this call to immediately after `const controller = new PanelController(session, prepared.bridge);` instead, and keep the `recordAuthDebug` call where it is.)

- [ ] **Step 7: Full gate + commit**

```bash
bun run typecheck
npx vitest run packages/web-shell/src/controller.test.ts
npx prettier --check packages/web-shell/src/controller.ts packages/web-shell/src/controller.test.ts packages/web-shell/src/taskpane/main.tsx
git add packages/web-shell/src/controller.ts packages/web-shell/src/controller.test.ts packages/web-shell/src/taskpane/main.tsx
git commit -m "feat(web-shell): expose discovered skills/data stores in PanelController state"
```

---

## Roadmap note (out of scope for this plan, tracked separately)

The following are explicitly DEFERRED, pending the still-running multi-team investigation into the live incident (context capture legibility, command-loop convergence, Office.js effectiveness, streaming resilience):
- Refactoring the EXISTING `outline`/`read`/`search`/`workspace`/`cat`/`grep` verbs to route through `DocFs` internally (unifying the two current code paths) — deferred because the investigation's context-engineering findings will likely prescribe specific `captureDocState`/outline-rendering changes, and doing this refactor first risks redoing it.
- Improving `bridge-excel`'s header/number-format capture legibility (the root cause of the blank-header/raw-time-serial incident).
- Improving the `<capabilities>` block's bulk-write guidance for large/irregular tables.
- Hardening the SSE/streaming path against proxy/tunnel chunk reordering.
- Feeding `availableDataStores` into an actual `@`-picker UI and `dataStoreSpecs` grounding call (Task 6 above stops at making the data reachable in `PanelController` state — rendering a picker and threading a selection into a turn's grounding is its own UI-plus-plumbing task, reasonably scoped separately once a design for the picker exists).

## Self-Review

- **Spec coverage:** `ls`/`find` verbs end-to-end (Tasks 1-2) ✔; skill catalog regeneration (Task 3) ✔; Python parity (Task 4) ✔; no-fence telemetry (Task 5) ✔; discovery reaching the UI layer (Task 6) ✔. Explicitly deferred items are listed above, not silently dropped.
- **Placeholder scan:** Tasks 1-3, 5-6 have concrete code throughout. Task 4's exact edit is intentionally contingent on Step 1's investigation (the real file/data-structure to modify is unknown until the implementer reads `parse_commands.py` and checks for a possibly-separate parity file) — this is a legitimate "read this first" scene-setting step, not a placeholder, because the plan is explicit about WHAT must be true afterward (both syntaxes accepted, both rejections preserved) even though the exact diff depends on code the plan's author hasn't fully inlined. Step 3 in Task 1/2's execution tests similarly asks the implementer to check the file's real test-construction pattern rather than guessing a fake API — also legitimate scene-setting, not a placeholder, since the assertions themselves are concrete.
- **Type consistency:** `ReadIntent`'s `ls`/`find` variants, `ParsedCommand`'s `ls`/`find` variants, and `compileCommand`'s mapping between them use identical field names (`path`, `glob`) across Tasks 1-2. `docFs`/`docFsLs`/`docFsFind` naming is consistent between Task 1's import statement and both tasks' execution-case code.
