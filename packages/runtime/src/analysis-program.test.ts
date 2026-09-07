import { describe, expect, it } from 'vitest';
import {
  AnalysisBindings,
  compileAnalysisProgram,
  inspectAnalysisProgram,
} from './analysis-program.js';

const A = `a_${'a'.repeat(24)}`;
const B = `a_${'b'.repeat(24)}`;
describe('typed artifact bindings', () => {
  it('resolves only artifact fields, preserving document targets and literal SQL strings', () => {
    const bindings = new AnalysisBindings();
    bindings.bind('source', A);
    expect(
      bindings.resolve({
        kind: 'query',
        inputs: ['$source'],
        sql: "SELECT '$source' AS literal, c0 FROM $source",
      }),
    ).toMatchObject({ inputs: [A], sql: `SELECT '$source' AS literal, c0 FROM ${A}` });
    expect(
      bindings.resolve({ kind: 'materialize', id: '$source', destination: '$source!A1' }),
    ).toEqual({ kind: 'materialize', id: A, destination: '$source!A1' });
    expect(bindings.resolve({ kind: 'capture', range: '$source!A1' })).toMatchObject({
      range: '$source!A1',
    });
  });
  it('rejects undeclared inputs, rebindings, non-artifact handles, and expired names', () => {
    const bindings = new AnalysisBindings();
    bindings.bind('source', A);
    bindings.bind('other', B);
    expect(() => bindings.bind('source', B)).toThrow('already exists');
    expect(() => bindings.bind('x', 'result:123')).toThrow('artifact');
    expect(() =>
      bindings.resolve({ kind: 'query', inputs: ['$source'], sql: 'SELECT * FROM $other' }),
    ).toThrow('inputs');
    expect(() =>
      bindings.resolve({ kind: 'query', inputs: ['$source'], sql: 'SELECT * FROM foo$source' }),
    ).toThrow('Invalid');
    expect(() =>
      bindings.resolve({
        kind: 'query',
        inputs: ['$source'],
        sql: 'SELECT * FROM $source; DELETE FROM x',
      }),
    ).toThrow('single analytical');
    bindings.clear();
    expect(() => bindings.resolve({ kind: 'inspect', id: '$source' })).toThrow('Unknown');
  });
  it('exposes detached binding entries for deterministic execution-state projection', () => {
    const bindings = new AnalysisBindings();
    bindings.bind('source', A);
    const entries = bindings.entries();
    bindings.clear();
    expect(entries).toEqual([['source', A]]);
    expect(bindings.entries()).toEqual([]);
  });
  it('does not reinterpret quotes, comments or SQL strings as executable bindings', () => {
    const bindings = new AnalysisBindings();
    bindings.bind('x', A);
    const resolve = (sql: string) => bindings.resolve({ kind: 'query', inputs: ['$x'], sql });
    expect(resolve("SELECT 'it''s $x' FROM $x")).toMatchObject({
      sql: `SELECT 'it''s $x' FROM ${A}`,
    });
    expect(() => resolve('SELECT * FROM "$x"')).toThrow();
    expect(() => resolve('SELECT * FROM $x -- $x')).toThrow();
    expect(() => resolve("SELECT 'unclosed $x")).toThrow();
  });
});

describe('SDK analysis program compiler', () => {
  it('preserves a typed nonempty guard and serializes dependency barriers around writes', () => {
    const program = {
      version: 1 as const,
      steps: [
        { op: 'bind' as const, name: 'source', action: { kind: 'capture' as const, range: 'A1' } },
        { op: 'materialize' as const, id: '$source', destination: 'B1', whenNonEmpty: true },
        { op: 'bind' as const, name: 'after', action: { kind: 'capture' as const, range: 'B1' } },
      ],
    };
    expect(compileAnalysisProgram(program)).toContain('"whenNonEmpty":true');
    expect(inspectAnalysisProgram(program).steps.map((step) => step.dependsOn)).toEqual([
      [],
      [0],
      [1],
    ]);
  });
  it('emits a bounded, locally executable program with verified completion', () => {
    const text = compileAnalysisProgram({
      version: 1,
      steps: [
        { op: 'bind', name: 'source', action: { kind: 'capture', range: 'S!A1:B2' } },
        {
          op: 'bind',
          name: 'sum',
          action: { kind: 'query', inputs: ['$source'], sql: 'SELECT sum(c1) FROM $source' },
        },
        { op: 'materialize', id: '$sum', destination: 'S!D1' },
      ],
    });
    expect(text).toContain('let $source = analyze ');
    expect(text).toContain('"id":"$sum"');
    expect(text.split('\n')).toHaveLength(4);
    expect(text.endsWith('finish when=verified')).toBe(true);
  });
  it('rejects forward references, duplicate names, non-read bindings and frame injection', () => {
    expect(() =>
      compileAnalysisProgram({
        version: 1,
        steps: [{ op: 'materialize', id: '$later', destination: 'D1' }],
      }),
    ).toThrow('Unbound');
    expect(() =>
      compileAnalysisProgram({
        version: 1,
        steps: [
          { op: 'bind', name: 'x', action: { kind: 'materialize', id: A, destination: 'D1' } },
        ],
      }),
    ).toThrow('artifact-producing');
    expect(() =>
      compileAnalysisProgram({
        version: 1,
        steps: [
          { op: 'bind', name: 'x', action: { kind: 'capture', range: '```\nset A1 hacked' } },
        ],
      }),
    ).toThrow('fence');
    expect(() =>
      compileAnalysisProgram({
        version: 1,
        steps: [
          { op: 'bind', name: 'x', action: { kind: 'capture', range: 'A1' } },
          { op: 'bind', name: 'x', action: { kind: 'capture', range: 'B1' } },
        ],
      }),
    ).toThrow('Duplicate');
  });
});
