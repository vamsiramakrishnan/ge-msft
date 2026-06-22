import { describe, it, expect } from 'vitest';
import {
  ParsedExprSchema,
  EFFECT_COMPOSE_ERROR,
  isExpressionLine,
  parseExpressionLine,
  isExprParseError,
} from './expr-grammar.js';
import { EFFECT_VERBS } from './expr-grammar.js';
import {
  parseProgramBlock,
  isProgramExpr,
  isProgramCommand,
  WRITE_VERB_TO_KIND,
} from './command-grammar.js';

describe('expr-grammar — effect-verb set stays in sync with the command grammar', () => {
  it('EFFECT_VERBS covers exactly the ADR-0004 write verbs', () => {
    expect([...EFFECT_VERBS].sort()).toEqual(Object.keys(WRITE_VERB_TO_KIND).sort());
  });
});

describe('expr-grammar — line classification', () => {
  it('flags lines with a top-level pipe or a leading let', () => {
    expect(isExpressionLine('read Sales!A1:C9 | sum amount')).toBe(true);
    expect(isExpressionLine('let $t = read X')).toBe(true);
    expect(isExpressionLine('$t | count')).toBe(true);
  });

  it('does NOT flag plain simple commands (no pipe, not let)', () => {
    expect(isExpressionLine('read Sales!A1:C9')).toBe(false);
    expect(isExpressionLine('outline')).toBe(false);
    expect(isExpressionLine('search revenue')).toBe(false);
    expect(isExpressionLine('set Sales!F2 =C2-D2')).toBe(false);
    // A pipe INSIDE a quoted arg is not a top-level pipe.
    expect(isExpressionLine('comment Sales!A1 "a | b"')).toBe(false);
  });
});

describe('expr-grammar — pipeline parsing', () => {
  it('parses read + filter + sum', () => {
    const expr = parseExpressionLine('read Sales!A1:C9 | filter region=East | sum amount');
    expect(isExprParseError(expr)).toBe(false);
    expect(expr).toEqual({
      kind: 'pipeline',
      source: { src: 'read', selector: 'Sales!A1:C9' },
      stages: [
        { name: 'filter', args: 'region=East' },
        { name: 'sum', args: 'amount' },
      ],
    });
  });

  it('parses each source kind', () => {
    expect(parseExpressionLine('search revenue | count')).toMatchObject({
      source: { src: 'search', text: 'revenue' },
    });
    expect(parseExpressionLine('outline | count')).toMatchObject({ source: { src: 'outline' } });
    expect(parseExpressionLine('$t | count')).toMatchObject({ source: { src: 'var', name: 't' } });
    // Empty read selector (whole-document) is valid.
    expect(parseExpressionLine('read | count')).toMatchObject({
      source: { src: 'read', selector: '' },
    });
  });

  it('parses a stage with no args (count)', () => {
    const expr = parseExpressionLine('read X | count');
    expect(expr).toMatchObject({ stages: [{ name: 'count', args: '' }] });
  });

  it('validates against the Zod schema', () => {
    const expr = parseExpressionLine('read X | sum amount');
    expect(ParsedExprSchema.safeParse(expr).success).toBe(true);
  });
});

describe('expr-grammar — let bindings', () => {
  it('parses let $name = <pipeline>', () => {
    const expr = parseExpressionLine('let $east = read Sales!A1:C9 | filter region=East');
    expect(expr).toEqual({
      kind: 'let',
      name: 'east',
      pipeline: {
        kind: 'pipeline',
        source: { src: 'read', selector: 'Sales!A1:C9' },
        stages: [{ name: 'filter', args: 'region=East' }],
      },
    });
    expect(ParsedExprSchema.safeParse(expr).success).toBe(true);
  });

  it('rejects a non-$ name, a missing =, an empty pipeline', () => {
    expect(parseExpressionLine('let east = read X')).toMatchObject({ error: expect.any(String) });
    expect(parseExpressionLine('let $east read X')).toMatchObject({ error: expect.any(String) });
    expect(parseExpressionLine('let $east =')).toMatchObject({ error: expect.any(String) });
  });
});

describe('expr-grammar — Phase-1 pure-only guard (effects)', () => {
  it('rejects an effect verb as a pipeline source with the corrective', () => {
    expect(parseExpressionLine('set Sales!F2 =1 | sum x')).toEqual({ error: EFFECT_COMPOSE_ERROR });
    expect(parseExpressionLine('suggest "a" => "b" | count')).toEqual({
      error: EFFECT_COMPOSE_ERROR,
    });
  });

  it('rejects an unknown source distinctly from an effect', () => {
    const err = parseExpressionLine('frobnicate X | count');
    expect(isExprParseError(err)).toBe(true);
    expect((err as { error: string }).error).toContain('unknown pipeline source');
  });
});

describe('parseProgramBlock — mixed simple commands + expressions', () => {
  it('routes a plain read line to the simple command parser, a pipeline to the expr parser', () => {
    const block = '```cmd\nread Sales!A1:C9\nread Sales!A1:C9 | sum amount\n```';
    const { found, entries } = parseProgramBlock(block);
    expect(found).toBe(true);
    expect(entries).toHaveLength(2);
    // A plain `read X` is the simple command, NOT a pipeline.
    expect(isProgramCommand(entries[0]!)).toBe(true);
    expect(entries[0]).toEqual({ verb: 'read', selector: 'Sales!A1:C9' });
    // The piped line is an expression.
    expect(isProgramExpr(entries[1]!)).toBe(true);
  });

  it('handles a set line + a pipeline line in one block', () => {
    const block =
      '```cmd\nset Sales!F2 =C2-D2\nlet $t = read Sales!A1:C9 | filter region=East\n```';
    const { entries } = parseProgramBlock(block);
    expect(entries[0]).toEqual({ verb: 'set', cell: 'Sales!F2', value: '=C2-D2' });
    expect(isProgramExpr(entries[1]!)).toBe(true);
    expect(entries[1]).toMatchObject({ kind: 'let', name: 't' });
  });

  it('preserves ADR-0004 behavior: no fence → not found', () => {
    expect(parseProgramBlock('no block here').found).toBe(false);
  });

  it('skips blank and comment lines', () => {
    const block = '```cmd\n# a comment\n\nread X | count\n```';
    const { entries } = parseProgramBlock(block);
    expect(entries).toHaveLength(1);
    expect(isProgramExpr(entries[0]!)).toBe(true);
  });
});
