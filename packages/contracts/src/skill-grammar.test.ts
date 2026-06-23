import { describe, it, expect } from 'vitest';
import { parseProgramBlock, isProgramSkillDef, isProgramSkillCall } from './command-grammar.js';
import {
  parseSkillDefHeader,
  parseSkillCall,
  shadowsBuiltin,
  isSkillParseError,
  type ParsedSkillDef,
  type ParsedSkillCall,
} from './skill-grammar.js';

/** Wrap raw program lines in a ```cmd fence for parseProgramBlock. */
function fence(...lines: string[]): string {
  return '```cmd\n' + lines.join('\n') + '\n```';
}

describe('skill-grammar — parseSkillDefHeader (ADR-0005 Phase 3)', () => {
  it('parses a header with params', () => {
    expect(parseSkillDefHeader('def reconcile($a $b):')).toEqual({
      name: 'reconcile',
      params: ['a', 'b'],
    });
  });

  it('parses a zero-param header', () => {
    expect(parseSkillDefHeader('def snapshot():')).toEqual({ name: 'snapshot', params: [] });
  });

  it('accepts comma-separated params and bare (non-$) names', () => {
    expect(parseSkillDefHeader('def f(a, b, c):')).toEqual({ name: 'f', params: ['a', 'b', 'c'] });
  });

  it('rejects a name that shadows a built-in verb', () => {
    const r = parseSkillDefHeader('def set($a):');
    expect(isSkillParseError(r) && r.error).toContain('shadows a built-in');
  });

  it('rejects a duplicate parameter', () => {
    const r = parseSkillDefHeader('def f($a $a):');
    expect(isSkillParseError(r) && r.error).toContain('duplicate');
  });

  it('rejects a missing colon / parens / bad name', () => {
    expect(isSkillParseError(parseSkillDefHeader('def f($a)'))).toBe(true); // no colon
    expect(isSkillParseError(parseSkillDefHeader('def f:'))).toBe(true); // no parens
    expect(isSkillParseError(parseSkillDefHeader('def 9bad($a):'))).toBe(true); // bad name
  });
});

describe('skill-grammar — shadowsBuiltin', () => {
  it('flags every built-in verb', () => {
    for (const v of ['set', 'suggest', 'comment', 'read', 'done', 'help', 'slide', 'post']) {
      expect(shadowsBuiltin(v)).toBe(true);
    }
  });
  it('does not flag a fresh name', () => {
    expect(shadowsBuiltin('reconcile')).toBe(false);
  });
});

describe('skill-grammar — parseSkillCall (quote-aware args)', () => {
  it('scans bare and quoted positional args', () => {
    expect(parseSkillCall('f', 'Sales!A1 "two words" bare')).toEqual({
      kind: 'skill-call',
      name: 'f',
      args: ['Sales!A1', 'two words', 'bare'],
    });
  });
  it('a zero-arg call yields an empty args array', () => {
    expect(parseSkillCall('f', '')).toEqual({ kind: 'skill-call', name: 'f', args: [] });
  });
  it('an unterminated quoted arg is corrective', () => {
    const r = parseSkillCall('f', '"open');
    expect(isSkillParseError(r) && r.error).toContain('unterminated');
  });
});

describe('parseProgramBlock — def … end grouping', () => {
  it('groups a def block into ONE ParsedSkillDef entry with verbatim body', () => {
    const { entries } = parseProgramBlock(
      fence('def reconcile($a $b):', '  read $a', '  set $b = ($a | sum amount)', 'end'),
    );
    expect(entries).toHaveLength(1);
    const def = entries[0]!;
    expect(isProgramSkillDef(def)).toBe(true);
    expect(def as ParsedSkillDef).toEqual({
      kind: 'skill-def',
      name: 'reconcile',
      params: ['a', 'b'],
      body: ['read $a', 'set $b = ($a | sum amount)'],
    });
  });

  it('skips blanks/comments inside a body', () => {
    const { entries } = parseProgramBlock(fence('def f($a):', '', '# a note', '  read $a', 'end'));
    expect((entries[0] as ParsedSkillDef).body).toEqual(['read $a']);
  });

  it('a registered skill name parses as a CALL, not an unknown verb', () => {
    const { entries } = parseProgramBlock(
      fence('reconcile Sales!A1 Sales!B1'),
      new Set(['reconcile']),
    );
    expect(entries).toHaveLength(1);
    expect(isProgramSkillCall(entries[0]!)).toBe(true);
    expect(entries[0] as ParsedSkillCall).toEqual({
      kind: 'skill-call',
      name: 'reconcile',
      args: ['Sales!A1', 'Sales!B1'],
    });
  });

  it('an UNregistered name still parses as an unknown-verb command error (back-compat)', () => {
    const { entries } = parseProgramBlock(fence('reconcile Sales!A1')); // no knownSkills
    expect(entries[0]).toMatchObject({ error: expect.stringContaining('unknown verb') });
  });

  it('a def and a following call in the same block both parse (grouping then call)', () => {
    const { entries } = parseProgramBlock(
      fence('def f($a):', '  read $a', 'end', 'f Sales!A1'),
      new Set(['f']),
    );
    expect(entries).toHaveLength(2);
    expect(isProgramSkillDef(entries[0]!)).toBe(true);
    expect(isProgramSkillCall(entries[1]!)).toBe(true);
  });

  it('an unterminated def (no end) is corrective, not a throw', () => {
    const { entries } = parseProgramBlock(fence('def f($a):', '  read $a'));
    expect(entries[0]).toMatchObject({ error: expect.stringContaining('missing a closing') });
  });

  it('a stray end (no def) is corrective', () => {
    const { entries } = parseProgramBlock(fence('end'));
    expect(entries[0]).toMatchObject({ error: expect.stringContaining('without a matching') });
  });

  it('a nested def is rejected', () => {
    const { entries } = parseProgramBlock(
      fence('def outer($a):', '  def inner($b):', '  end', 'end'),
    );
    expect(entries[0]).toMatchObject({ error: expect.stringContaining('nested def') });
  });

  it('a malformed header is corrective but still consumes the body up to end', () => {
    const { entries } = parseProgramBlock(fence('def set($a):', '  read $a', 'end', 'read Foo'));
    expect(entries).toHaveLength(2);
    expect(entries[0]).toMatchObject({ error: expect.stringContaining('shadows a built-in') });
    expect(entries[1]).toMatchObject({ verb: 'read' });
  });

  it('all ADR-0004 lines parse unchanged when there are no skills/defs', () => {
    const { entries } = parseProgramBlock(fence('read Foo', 'set A1 1', 'done'));
    expect(entries.map((e) => (e as { verb?: string }).verb)).toEqual(['read', 'set', 'done']);
  });
});
