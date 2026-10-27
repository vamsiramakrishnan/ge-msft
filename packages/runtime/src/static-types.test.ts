import { describe, expect, it } from 'vitest';
import { parseProgramBlock } from '@ge/contracts';
import { typecheckProgram, inferPipeline, type StaticBinding } from './static-types.js';

const program = (text: string) => parseProgramBlock('```cmd\n' + text + '\n```').entries;
const check = (text: string, scope: Array<[string, StaticBinding]> = []) =>
  typecheckProgram(program(text), new Map(scope)).diagnostics.map((d) => d.message);

describe('static type pass (B2)', () => {
  it('accepts a well-typed program and binds the inferred types', () => {
    const entries = program(
      'let $t = table [["region","amount"],["East","5"]]\nlet $n = $t | filter region=East | sum amount\nset A1 = ($n)',
    );
    const { diagnostics, env } = typecheckProgram(entries, new Map());
    expect(diagnostics).toEqual([]);
    expect(env.get('t')).toEqual({ type: 'Table', columns: ['region', 'amount'] });
    expect(env.get('n')).toEqual({ type: 'Number' });
  });

  it('rejects an unknown transform with a did-you-mean', () => {
    expect(check('let $x = read A1:B3 | sumx amount')).toEqual([
      'unknown transform "sumx" — did you mean "sum"?',
    ]);
  });

  it('rejects a transform fed a number (sum | head)', () => {
    expect(check('let $x = read A1:B3 | sum amount | head 3')).toEqual([
      '"head" needs a table, but the value before it is a number',
    ]);
  });

  it('rejects an unbound $name, including inside an effect expression and an each header', () => {
    expect(check('set A1 = ($ghost | sum x)')).toEqual([
      '$ghost is an unbound variable (used before it is bound)',
    ]);
    expect(check('each $r in $rows:\n  set A1 $r.x\nend')).toEqual([
      '$rows is an unbound variable (used before it is bound)',
    ]);
  });

  it('accepts a binding the running loop already holds', () => {
    expect(check('let $y = $t | sort amount', [['t', { type: 'Table' }]])).toEqual([]);
  });

  it('applies column laws only when the columns are statically known', () => {
    expect(check('let $t = table [["a","b"],["1","2"]]\nlet $u = $t | select a, zz')).toEqual([
      '"select" names column "zz" the table does not have (columns: a, b)',
    ]);
    // A host read's columns are unknown: no law is applied.
    expect(check('let $u = read A1:B3 | select zz')).toEqual([]);
    // Columns follow select / rename / derive / group.
    expect(
      check(
        'let $t = table [["a","b"],["1","2"]]\nlet $u = $t | rename a=x | derive c = x * 2 | select x, c',
      ),
    ).toEqual([]);
    expect(
      check(
        'let $t = table [["a","b"],["1","2"]]\nlet $u = $t | group a agg sum(b) as total | sort nope',
      ),
    ).toEqual(['"sort" names column "nope" the table does not have (columns: a, total)']);
  });

  it('checks each-body row references against known columns', () => {
    expect(
      check('let $t = table [["name","qty"],["x","1"]]\neach $r in $t:\n  set A1 $r.price\nend'),
    ).toEqual(['each body names column "price" that $t does not have (columns: name, qty)']);
  });

  it('refuses an analysis artifact in a pipeline and reports one diagnostic per entry', () => {
    const entries = program(
      'let $a = analyze {"kind":"capture","range":"A1:B2"}\nlet $b = $a | sum x\nlet $c = $zz | count',
    );
    const { diagnostics } = typecheckProgram(entries, new Map());
    expect(diagnostics).toEqual([
      { index: 1, message: '$a is an analysis artifact; use analyze to read it' },
      { index: 2, message: '$zz is an unbound variable (used before it is bound)' },
    ]);
  });

  it('types a join against a bound table and refuses a non-table partner', () => {
    const scope = new Map<string, StaticBinding>([['n', { type: 'Number' }]]);
    const entries = program('let $j = read A1:B2 | join $n on id');
    expect(typecheckProgram(entries, scope).diagnostics[0]?.message).toBe(
      'join needs a bound table, but $n is a number',
    );
    const ok = program(
      'let $t = table [["id","v"],["1","2"]]\nlet $j = read A1:B2 | join $t on id',
    );
    expect(typecheckProgram(ok, new Map()).diagnostics).toEqual([]);
    expect(inferPipeline(program('read A1 | count')[0] as never, new Map()).binding).toEqual({
      type: 'Number',
    });
  });
});
