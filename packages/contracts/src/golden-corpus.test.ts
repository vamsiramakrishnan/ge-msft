import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  extractCommandBlock,
  isCommandParseError,
  parseCommandLine,
  parseProgramBlock,
  type CommandParseError,
  type ProgramEntry,
} from './command-grammar.js';

const isErrorEntry = (e: ProgramEntry): e is CommandParseError => 'error' in e;

/**
 * The differential half of the ADR-0008 §4 conformance gate. `skill/golden-corpus.jsonl` pins the
 * behaviours BOTH parsers must agree on; `skill/parity_corpus_test.py` drives it through the Python
 * preflight, and this file drives the very same cases, with the very same expectation semantics,
 * through the authoritative TypeScript grammar. A case that passes on one side and fails on the
 * other is a grammar drift, caught here before it reaches a live turn.
 *
 * Expectation shapes (mirroring `_check_flat` / `_check_block` in the Python harness):
 *   • flat (no fence, no block-level key): `verbs: [v]` → the line parses to verb `v`;
 *     `verbs: []` → the line is skipped (blank/comment); `error_substring` → an error containing it.
 *   • block (a ```cmd fence or `complete`/`no_block`/`has_error`): `verbs` is the ordered list of
 *     command verbs (expressions/bindings carry no verb), `has_error`, `complete` (closed fence +
 *     no errors + a `done`), `command_kind`/`argument_text` for a single `/<kind>` invoke.
 */

interface Expectation {
  verbs?: string[];
  error_substring?: string;
  has_error?: boolean;
  complete?: boolean;
  no_block?: boolean;
  command_kind?: string;
  argument_text?: string;
}
interface Case {
  name: string;
  input: string;
  expect: Expectation;
}

const CORPUS = resolve(
  dirname(fileURLToPath(import.meta.url)),
  '../../../skill/golden-corpus.jsonl',
);
const BLOCK_KEYS = ['complete', 'no_block', 'has_error'] as const;

function loadCases(): Case[] {
  return readFileSync(CORPUS, 'utf8')
    .split('\n')
    .map((raw) => raw.trim())
    .filter((raw) => raw !== '')
    .map((raw) => JSON.parse(raw) as Case);
}

function isBlockCase(c: Case): boolean {
  return c.input.includes('```cmd') || BLOCK_KEYS.some((k) => k in c.expect);
}

/** The Python harness's `block_is_complete`: closed fence, no errors, at least one `done`. */
function blockIsComplete(modelText: string, entries: ProgramEntry[]): boolean {
  const fences = modelText
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l.startsWith('```'));
  const closed = fences.length >= 2;
  if (!closed) return false;
  if (entries.some(isErrorEntry)) return false;
  return entries.some((e) => 'verb' in e && e.verb === 'done');
}

function verbsOf(entries: ProgramEntry[]): string[] {
  return entries.flatMap((e) => ('verb' in e ? [e.verb] : []));
}

describe('golden corpus — TypeScript side of the TS≡Python conformance gate', () => {
  const cases = loadCases();

  it('loads a non-trivial corpus', () => {
    expect(cases.length).toBeGreaterThan(50);
  });

  for (const c of cases) {
    it(`[${c.name}] ${isBlockCase(c) ? 'block' : 'flat'}`, () => {
      if (!isBlockCase(c)) {
        const rec = parseCommandLine(c.input);
        if (c.expect.error_substring !== undefined) {
          expect(isCommandParseError(rec)).toBe(true);
          expect(isCommandParseError(rec) && rec.error).toContain(c.expect.error_substring);
          return;
        }
        const verbs = c.expect.verbs ?? [];
        if (verbs.length === 0) {
          // A skipped line: the line parser reports "empty command line" where Python yields None.
          expect(isCommandParseError(rec) && rec.error).toBe('empty command line');
          return;
        }
        expect(isCommandParseError(rec)).toBe(false);
        expect('verb' in rec && rec.verb).toBe(verbs[0]);
        return;
      }

      if (c.expect.no_block) {
        expect(extractCommandBlock(c.input)).toBeNull();
        return;
      }
      const { entries } = parseProgramBlock(c.input);
      const hasError = entries.some(isErrorEntry);
      if (c.expect.verbs !== undefined) expect(verbsOf(entries)).toEqual(c.expect.verbs);
      if (c.expect.has_error !== undefined) expect(hasError).toBe(c.expect.has_error);
      if (c.expect.command_kind !== undefined) {
        const cmds = entries.filter((e) => 'verb' in e);
        expect(cmds).toHaveLength(1);
        const only = cmds[0]!;
        expect('verb' in only && only.verb).toBe('invoke');
        if ('verb' in only && only.verb === 'invoke') {
          expect(only.kind).toBe(c.expect.command_kind);
          expect(only.props['text']).toBe(c.expect.argument_text);
        }
      }
      if (c.expect.complete !== undefined) {
        const complete = blockIsComplete(c.input, entries);
        expect(complete).toBe(c.expect.complete);
        // Fail-closed invariant: a block with any parse error is never complete.
        expect(hasError && complete).toBe(false);
      }
    });
  }

  it('annotates block-level parse errors with the fence-relative line and source', () => {
    const { entries } = parseProgramBlock('```cmd\noutline\n\n# note\nwrit A1 5\ndone\n```');
    const errors = entries.filter(isErrorEntry);
    expect(errors).toHaveLength(1);
    expect(errors[0]).toMatchObject({ line: 4, source: 'writ A1 5' });
    expect(errors[0]!.error).toContain('did you mean "grid"');
  });
});
