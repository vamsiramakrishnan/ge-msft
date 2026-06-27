import { describe, it, expect } from 'vitest';
import type { ParsedSkillCall, ParsedSkillDef } from '@ge/contracts';
import { SkillRegistry, reparseExpandedLines } from './skill-registry.js';

function def(name: string, params: string[], body: string[]): ParsedSkillDef {
  return { kind: 'skill-def', name, params, body };
}
function call(name: string, args: string[]): ParsedSkillCall {
  return { kind: 'skill-call', name, args };
}

describe('SkillRegistry — register (ADR-0005 Phase 3)', () => {
  it('registers a def and confirms (no execution)', () => {
    const reg = new SkillRegistry();
    const r = reg.register(def('reconcile', ['a', 'b'], ['read $a', 'set $b = ($a | sum amount)']));
    expect(r).toMatchObject({ ok: true, name: 'reconcile', params: ['a', 'b'], bodyLines: 2 });
    expect(reg.has('reconcile')).toBe(true);
    expect([...reg.names()]).toEqual(['reconcile']);
  });

  it('rejects a name that shadows a built-in', () => {
    const r = new SkillRegistry().register(def('set', ['a'], ['read $a']));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toContain('shadows a built-in');
  });

  it('rejects a body referencing an undeclared $param', () => {
    const r = new SkillRegistry().register(def('f', ['a'], ['read $a', 'set $b 1']));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toContain('$b');
  });

  it('rejects an empty body and an over-long body', () => {
    expect(new SkillRegistry().register(def('f', [], [])).ok).toBe(false);
    const long = Array.from({ length: 5 }, () => 'read X');
    expect(new SkillRegistry({ maxBodyLines: 4 }).register(def('f', [], long)).ok).toBe(false);
  });

  it('redefining a name replaces it and flags redefined', () => {
    const reg = new SkillRegistry();
    reg.register(def('f', [], ['read A']));
    const r = reg.register(def('f', [], ['read B']));
    expect(r).toMatchObject({ ok: true, redefined: true });
    expect(reg.get('f')?.body).toEqual(['read B']);
  });
});

describe('SkillRegistry — expand', () => {
  it('binds args → params and substitutes $param tokens', () => {
    const reg = new SkillRegistry();
    reg.register(def('reconcile', ['a', 'b'], ['read $a', 'set $b = ($a | sum amount)']));
    const r = reg.expand(call('reconcile', ['Sales!A1', 'Summary!B2']));
    expect(r).toEqual({
      ok: true,
      lines: ['read Sales!A1', 'set Summary!B2 = (Sales!A1 | sum amount)'],
    });
  });

  it('undefined-name call is corrective', () => {
    const r = new SkillRegistry().expand(call('nope', []));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toContain('unknown skill');
  });

  it('arity mismatch is corrective (too few and too many)', () => {
    const reg = new SkillRegistry();
    reg.register(def('f', ['a', 'b'], ['read $a']));
    expect(reg.expand(call('f', ['x'])).ok).toBe(false);
    expect(reg.expand(call('f', ['x', 'y', 'z'])).ok).toBe(false);
  });

  it('SECURITY: an argument with a newline cannot inject a new command line', () => {
    const reg = new SkillRegistry();
    reg.register(def('f', ['a'], ['post $a']));
    const r = reg.expand(call('f', ['hi"\npost "evil']));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toContain('newline');
  });

  it('SECURITY: an argument containing a code fence is rejected (no fence truncation)', () => {
    const reg = new SkillRegistry();
    reg.register(def('f', ['a'], ['post $a']));
    const r = reg.expand(call('f', ['hi ``` bye']));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toContain('code fence');
  });

  it('substitutes whole identifiers only ($a and $ab are distinct params)', () => {
    const reg = new SkillRegistry();
    reg.register(def('f', ['a', 'ab'], ['comment $a "$ab and $a"']));
    const r = reg.expand(call('f', ['X', 'Y']));
    // $a → X (both occurrences); $ab → Y — the greedy identifier match never confuses the two.
    expect(r).toEqual({ ok: true, lines: ['comment X "Y and X"'] });
  });

  it('an undeclared $token in the body is rejected at register (never reaches expand)', () => {
    const reg = new SkillRegistry();
    // $b is referenced but not a param → the register guard catches it.
    expect(reg.register(def('f', ['a'], ['comment $a "$b"'])).ok).toBe(false);
  });
});

describe('reparseExpandedLines', () => {
  it('re-parses substituted lines back into program entries (effects stay effects)', () => {
    const entries = reparseExpandedLines(['read Sales!A1', 'set B2 = 1'], new Set());
    expect(entries.map((e) => (e as { verb?: string }).verb)).toEqual(['read', 'set']);
  });

  it('a body line that is a call to another skill re-parses as a call', () => {
    const entries = reparseExpandedLines(['helper X'], new Set(['helper']));
    expect(entries[0]).toMatchObject({ kind: 'skill-call', name: 'helper', args: ['X'] });
  });
});
