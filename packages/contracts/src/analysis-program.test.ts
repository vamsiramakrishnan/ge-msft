import { describe, expect, it } from 'vitest';
import {
  isProgramAnalysisBinding,
  isProgramCommand,
  isProgramExpr,
  isProgramVerifiedFinish,
  parseProgramBlock,
} from './command-grammar.js';
import { parseSkillDefHeader } from './skill-grammar.js';

const parse = (...lines: string[]) =>
  parseProgramBlock(['```cmd', ...lines, '```'].join('\n')).entries;

describe('artifact bindings and verified completion', () => {
  it('retains artifact references in typed bindings without rewriting arbitrary strings', () => {
    const requests = [
      { kind: 'capture', range: 'Invoices!A1:D100' },
      { kind: 'query', inputs: ['$source'], sql: "SELECT '$source' FROM source" },
      { kind: 'reconcile', spec: { left: '$source', right: '$payments' } },
      { kind: 'filter', id: '$result', status: 'variance' },
      { kind: 'inspect', id: '$result' },
    ];
    for (const request of requests) {
      const entries = parse(`let $_result1 = analyze ${JSON.stringify(request)}`);
      expect(entries).toEqual([
        { kind: 'analysis-binding', name: '_result1', request: JSON.stringify(request) },
      ]);
      expect(isProgramAnalysisBinding(entries[0]!)).toBe(true);
      expect(isProgramExpr(entries[0]!)).toBe(false);
      expect(isProgramCommand(entries[0]!)).toBe(false);
    }
  });

  it.each(['materialize', 'remove', 'recovery', 'undo', 'resume', 'forget', 'unknown', null, []])(
    'rejects a binding that cannot produce an artifact: %j',
    (kind) => {
      expect(parse(`let $x = analyze ${JSON.stringify({ kind })}`)[0]).toHaveProperty('error');
    },
  );

  it.each([
    'let x = analyze {"kind":"capture"}',
    'let $1bad = analyze {"kind":"capture"}',
    'let $two names = analyze {"kind":"capture"}',
    'let $hyphen-name = analyze {"kind":"capture"}',
    `let $${'x'.repeat(65)} = analyze {"kind":"capture"}`,
    'let $x = analyze',
    'let $x = analyze []',
    'let $x = analyze null',
    'let $x = analyze {"kind":"query","value":NaN}',
    'let $x = analyze {"kind":"capture"} extra',
    'let $x = analyze {"kind":"capture"} | head 1',
    `let $x = analyze ${JSON.stringify({ kind: 'capture', range: 'x'.repeat(32768) })}`,
  ])('rejects malformed or unbounded bindings', (line) => {
    expect(parse(line)[0]).toHaveProperty('error');
  });

  it('keeps existing expressions, direct analyze commands, and legacy done unchanged', () => {
    const entries = parse('let $x = read S!A1:B2', 'analyze {"kind":"inspect","id":"a_1"}', 'done');
    expect(entries[0]).toMatchObject({ kind: 'let', name: 'x' });
    expect(entries[1]).toEqual({ verb: 'analyze', request: '{"kind":"inspect","id":"a_1"}' });
    expect(entries[2]).toEqual({ verb: 'done' });
  });

  it('accepts a terminal completion request after a write without claiming it is verified', () => {
    const entries = parse('set S!A1 5', 'finish when=verified', '', '# trailing comment');
    expect(entries).toEqual([
      { verb: 'set', cell: 'S!A1', value: '5' },
      { kind: 'verified-finish' },
    ]);
    expect(isProgramVerifiedFinish(entries[1]!)).toBe(true);
    expect(isProgramCommand(entries[1]!)).toBe(false);
  });

  it.each([
    'finish',
    'finish now',
    'finish when=applied',
    'finish when=verified extra',
    'finish when=verified;',
  ])('rejects malformed completion', (line) => {
    expect(parse(line)[0]).toHaveProperty('error');
  });

  it.each(['done', 'set S!A1 6', 'finish when=verified', 'def example():'])(
    'rejects content after verified completion',
    (line) => {
      expect(parse('finish when=verified', line)).toContainEqual({
        error: 'finish when=verified must be the final program entry',
      });
    },
  );

  it('rejects completion in a truncated fence or plain marker while preserving legacy recovery', () => {
    for (const marker of ['```cmd', 'cmd']) {
      expect(parseProgramBlock(`${marker}\nfinish when=verified`).entries).toContainEqual({
        error: 'finish when=verified requires a closed cmd fence',
      });
      expect(parseProgramBlock(`${marker}\nread S!A1`).entries).toEqual([
        { verb: 'read', selector: 'S!A1' },
      ]);
    }
  });

  it.each([
    ['set S!A1 5', 'finish now'],
    ['set S!A1 5', 'finish when=verified', 'set S!A2 6'],
    ['set S!A1 5', 'malformed command', 'finish when=verified'],
  ])('never retains prefix writes when verified completion is malformed', (...lines) => {
    expect(parse(...lines).every((entry) => 'error' in entry)).toBe(true);
  });

  it('never retains prefix writes when verified completion is truncated', () => {
    expect(parseProgramBlock('```cmd\nset S!A1 5\nfinish when=verified').entries).toEqual([
      { error: 'finish when=verified requires a closed cmd fence' },
    ]);
  });

  it.each([
    '```cmd\nset S!A1 5\nfinish when=verified\n```\n```cmd\nset S!A2 6\n```',
    '```cmd\nset S!A1 5\n```\n```cmd\nset S!A2 6\nfinish when=verified\n```',
    'Leading explanation\n```cmd\nset S!A1 5\nfinish when=verified\n```',
    '```cmd\nset S!A1 5\nfinish when=verified\n```\nTrailing explanation',
    '```cmd\nset S!A1 5\nfinish when=verified\n```\n# trailing comment',
    '```cmd\nanalyze {"kind":"query","inputs":["a_id"],"sql":"SELECT \'```\' FROM a_id"}\nfinish when=verified\n```',
  ])('rejects ambiguous verified response framing without retaining prefix effects', (response) => {
    const result = parseProgramBlock(response);
    expect(result.found).toBe(true);
    expect(result.entries).toEqual([
      {
        error:
          'finish when=verified requires exactly one cmd fence with no surrounding text or embedded fences',
      },
    ]);
  });

  it('accepts whitespace around one complete verified response', () => {
    expect(
      parseProgramBlock(' \n```cmd\r\nset S!A1 5\r\nfinish when=verified\r\n```\n ').entries,
    ).toEqual([{ verb: 'set', cell: 'S!A1', value: '5' }, { kind: 'verified-finish' }]);
  });

  it('preserves legacy first-frame extraction without verified completion', () => {
    expect(
      parseProgramBlock(
        'Explanation\n```cmd\nset S!A1 5\n```\n```cmd\nset S!A2 6\n```\nTrailing explanation',
      ).entries,
    ).toEqual([{ verb: 'set', cell: 'S!A1', value: '5' }]);
  });

  it('reserves finish against skills, including previously registered names', () => {
    expect(parseSkillDefHeader('def finish():')).toHaveProperty('error');
    expect(parseSkillDefHeader('def FINISH():')).toHaveProperty('error');
    expect(
      parseProgramBlock('```cmd\nfinish when=verified\n```', new Set(['finish'])).entries,
    ).toEqual([{ kind: 'verified-finish' }]);
  });
});
