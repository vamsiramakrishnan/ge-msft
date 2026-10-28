import { describe, it, expect } from 'vitest';
import {
  ParsedEachBlockSchema,
  eachBodyLineError,
  isEachHeader,
  parseEachHeader,
} from './each-grammar.js';
import { isProgramEachBlock, parseProgramBlock, type ProgramEntry } from './command-grammar.js';
import { shadowsBuiltin } from './skill-grammar.js';

const errorsOf = (entries: ProgramEntry[]) =>
  entries.flatMap((e) => ('error' in e ? [e.error] : []));

describe('each-grammar — header', () => {
  it('parses `each $row in $table:`', () => {
    expect(parseEachHeader('each $r in $rows:')).toEqual({ rowVar: 'r', tableVar: 'rows' });
    expect(parseEachHeader('EACH $r IN $rows :')).toEqual({ rowVar: 'r', tableVar: 'rows' });
  });

  it('refuses a malformed header or identical names', () => {
    expect(parseEachHeader('each r in $rows:')).toMatchObject({
      error: expect.stringContaining('usage'),
    });
    expect(parseEachHeader('each $r in $rows')).toMatchObject({
      error: expect.stringContaining('usage'),
    });
    expect(parseEachHeader('each $r in $r:')).toMatchObject({
      error: expect.stringContaining('distinct'),
    });
  });

  it('isEachHeader is a loose head check; `each` is a structural keyword', () => {
    expect(isEachHeader('each $r in $t:')).toBe(true);
    expect(isEachHeader('eachother')).toBe(false);
    expect(shadowsBuiltin('each')).toBe(true);
  });
});

describe('each-grammar — body rules', () => {
  it('accepts effect lines that reference the row variable by column', () => {
    expect(eachBodyLineError('r', 'set Summary!B2 = $r.amount')).toBeUndefined();
    expect(eachBodyLineError('r', 'comment $r.cell "flag $r.note"')).toBeUndefined();
    expect(eachBodyLineError('r', 'set B2 = $total')).toBeUndefined(); // other $vars are fine
  });

  it('refuses non-effect lines, foreign row variables and a bare row variable', () => {
    expect(eachBodyLineError('r', 'read A1')).toContain('effect commands');
    expect(eachBodyLineError('r', 'let $x = $r.amount')).toContain('effect commands');
    expect(eachBodyLineError('r', 'set B2 = $row.amount')).toContain('unknown row variable $row');
    expect(eachBodyLineError('r', 'set B2 = $r')).toContain('bare $r');
  });
});

describe('each-grammar — block grouping in parseProgramBlock', () => {
  it('groups the block into one entry with verbatim body lines', () => {
    const { entries } = parseProgramBlock(
      '```cmd\nlet $t = read Sales!A1:B4\neach $r in $t:\n  set Summary!B2 = $r.amount\n  # note\n\n  comment $r.cell "x"\nend\ndone\n```',
    );
    expect(entries).toHaveLength(3);
    const each = entries[1]!;
    expect(isProgramEachBlock(each)).toBe(true);
    expect(each).toEqual({
      kind: 'each-block',
      rowVar: 'r',
      tableVar: 't',
      body: ['set Summary!B2 = $r.amount', 'comment $r.cell "x"'],
    });
    expect(ParsedEachBlockSchema.safeParse(each).success).toBe(true);
    expect(entries[2]).toEqual({ verb: 'done' });
  });

  it('reports a missing end, an empty body, a nested block and a non-effect line (with line numbers)', () => {
    expect(
      errorsOf(parseProgramBlock('```cmd\neach $r in $t:\nset B2 = $r.a\n```').entries),
    ).toEqual(['each $r in $t is missing a closing `end`']);
    expect(errorsOf(parseProgramBlock('```cmd\neach $r in $t:\nend\n```').entries)).toEqual([
      'each $r in $t has an empty body',
    ]);
    expect(
      errorsOf(
        parseProgramBlock('```cmd\neach $r in $t:\neach $s in $u:\nset B2 = 1\nend\nend\n```')
          .entries,
      ),
    ).toEqual([
      'nested each inside each $r is not allowed',
      '`end` without a matching `def` or `each`',
    ]);
    const bad = parseProgramBlock('```cmd\noutline\neach $r in $t:\nread $r.cell\nend\ndone\n```');
    const errors = bad.entries.filter((e) => 'error' in e);
    expect(errors).toHaveLength(1);
    expect(errors[0]).toMatchObject({ line: 2, source: 'each $r in $t:' });
    expect(errors[0]!.error).toContain('effect commands');
    // Recovery: the rest of the program still parses.
    expect(bad.entries.at(-1)).toEqual({ verb: 'done' });
  });

  it('a malformed header skips to end and reports the usage', () => {
    const { entries } = parseProgramBlock('```cmd\neach r in $t:\nset B2 = 1\nend\ndone\n```');
    expect(errorsOf(entries)).toEqual([expect.stringContaining('malformed each header')]);
    expect(entries.at(-1)).toEqual({ verb: 'done' });
  });
});
