import { z } from 'zod';
import { ALL_VERBS, scanQuoted } from './command-grammar.js';

/**
 * ADR-0005 Phase 3 — named skills: saved, parameterized compositions the model defines once and
 * calls. This module owns the **grammar + AST** for a skill *definition* (`def … end`) and a skill
 * *call* (`<name> <arg> …`); the runtime owns the registry, argument binding, body substitution,
 * and the expansion-into-plan (the existing Phase-2 plan machinery — a skill call introduces NO new
 * gate/approval bypass).
 *
 * SCOPE (this wave): **parameterized macros** — a `def` body is a sequence of normal command /
 * composition lines that may reference `$p1 … $pN`; a call binds positional args → params and
 * textually substitutes the `$param` tokens, producing expanded entries that run through the
 * Phase-2 plan (type-check → dry-run → approvePlan → gated execute). There is NO `for`/`each`
 * iteration yet (a later phase).
 *
 * The header grammar is:
 *   def <name>(<p1> <p2> …):   ← starts a definition; `<name>` is the skill, `<pi>` its params
 *     <body line>              ← normal command / composition lines, may reference `$pi`
 *     …
 *   end                        ← terminates the body
 *
 * A skill name may NOT shadow a built-in verb (`def set(...)` is rejected) — so a registered skill
 * never silently overrides a primitive, and the call dispatch stays unambiguous.
 */

/* ───────────────────────────── AST ───────────────────────────── */

/**
 * A parsed skill definition (the whole `def … end` block, grouped by {@link parseProgramBlock}).
 * `params` are the declared parameter names (WITHOUT the `$`); `body` are the VERBATIM body lines
 * (already stripped of indentation, blanks/comments removed) — the runtime re-parses each line
 * AFTER `$param` substitution, so the contracts layer never type-checks a half-bound body.
 */
export interface ParsedSkillDef {
  kind: 'skill-def';
  name: string;
  params: string[];
  body: string[];
}

/**
 * A parsed skill call (`<name> <arg1> <arg2> …`). The contracts parser emits this only when the
 * caller passed `<name>` in the `knownSkills` set (the runtime's registry) — so the parser stays
 * the single source of truth for *shape* while the runtime owns *which names are registered*. Args
 * are positional, quote-aware (a quoted arg may carry spaces); the runtime checks arity and binds
 * `args[i]` → `params[i]`.
 */
export interface ParsedSkillCall {
  kind: 'skill-call';
  name: string;
  args: string[];
}

/* ───────────────────────────── Zod ───────────────────────────── */

export const ParsedSkillDefSchema: z.ZodType<ParsedSkillDef> = z.object({
  kind: z.literal('skill-def'),
  name: z.string(),
  params: z.array(z.string()),
  body: z.array(z.string()),
});

export const ParsedSkillCallSchema: z.ZodType<ParsedSkillCall> = z.object({
  kind: z.literal('skill-call'),
  name: z.string(),
  args: z.array(z.string()),
});

export function isParsedSkillDef(node: { kind?: string }): node is ParsedSkillDef {
  return node.kind === 'skill-def';
}

export function isParsedSkillCall(node: { kind?: string }): node is ParsedSkillCall {
  return node.kind === 'skill-call';
}

/** A skill parse error carries a CLI-style corrective message (same shape as the other parsers). */
export interface SkillParseError {
  error: string;
}

export function isSkillParseError(node: object): node is SkillParseError {
  return 'error' in node && typeof (node as { error: unknown }).error === 'string';
}

/* ───────────────────────── line classification ─────────────────────────── */

/** The header line that opens a definition (`def name(...)`: ). Case-insensitive on the keyword. */
export function isSkillDefHeader(line: string): boolean {
  return /^def\s/i.test(line.trim());
}

/** The terminator line of a definition body. */
export function isSkillEnd(line: string): boolean {
  return line.trim().toLowerCase() === 'end';
}

/* ───────────────────────── call parsing ─────────────────────────── */

/**
 * Parse a skill-call line `<name> <arg1> <arg2> …` into a {@link ParsedSkillCall}. The caller has
 * already established `<name>` is a registered skill; here we scan the positional args. Each arg is
 * EITHER a double-quoted string (quote-aware via {@link scanQuoted}, so an arg may carry spaces and
 * is the substituted token verbatim) OR a bare whitespace-delimited token. A malformed quoted arg
 * (unterminated) is a corrective error. Zero args is valid (arity is checked by the runtime).
 */
export function parseSkillCall(name: string, rest: string): ParsedSkillCall | SkillParseError {
  const args: string[] = [];
  let i = 0;
  while (i < rest.length) {
    while (i < rest.length && /\s/.test(rest[i]!)) i++;
    if (i >= rest.length) break;
    if (rest[i] === '"') {
      const scanned = scanQuoted(rest, i);
      if (!scanned) return { error: `unterminated quoted argument in call to "${name}"` };
      args.push(scanned.value);
      i = scanned.end;
    } else {
      const sp = rest.slice(i).search(/\s/);
      const end = sp === -1 ? rest.length : i + sp;
      args.push(rest.slice(i, end));
      i = end;
    }
  }
  return { kind: 'skill-call', name, args };
}

/* ───────────────────────── identifier rules ─────────────────────────── */

/** A skill name: a plain identifier (letters/digits/underscore, not starting with a digit). */
const SKILL_NAME_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

/** A parameter name inside the header — a plain identifier OR a `$`-prefixed one (`$a` / `a`). */
function normalizeParam(token: string): string | undefined {
  const name = token.startsWith('$') ? token.slice(1) : token;
  return SKILL_NAME_RE.test(name) ? name : undefined;
}

/** True when `name` collides with a built-in verb (so a skill can never shadow a primitive). */
export function shadowsBuiltin(name: string): boolean {
  return (ALL_VERBS as readonly string[]).includes(name);
}

/* ───────────────────────── header parsing ─────────────────────────── */

/**
 * Parse a `def <name>(<p1> <p2> …):` header into `{ name, params }` or a corrective error. Params
 * are whitespace- and/or comma-separated, each a plain or `$`-prefixed identifier; a zero-param
 * skill is `def name():`. Rejects a name that shadows a built-in verb, a malformed name, a missing
 * `(…)` / trailing `:`, a bad param token, or a duplicate param.
 */
export function parseSkillDefHeader(
  line: string,
): { name: string; params: string[] } | SkillParseError {
  const usage = 'def needs a header — usage: def <name>($p1 $p2 …): then body lines, then end';
  const trimmed = line.trim();
  // Strip the leading `def` keyword.
  const afterDef = trimmed.slice(trimmed.search(/\s/) + 1).trim();

  const open = afterDef.indexOf('(');
  const close = afterDef.indexOf(')');
  if (open === -1 || close === -1 || close < open) return { error: usage };

  const name = afterDef.slice(0, open).trim();
  if (name === '') return { error: 'def needs a skill name — usage: def <name>(...):' };
  if (!SKILL_NAME_RE.test(name)) {
    return { error: `invalid skill name "${name}" — use letters, digits, and underscore` };
  }
  if (shadowsBuiltin(name)) {
    return { error: `skill name "${name}" shadows a built-in command — choose another name` };
  }

  // Everything after the `)` must be a (possibly empty before the colon) → exactly a trailing `:`.
  const tail = afterDef.slice(close + 1).trim();
  if (tail !== ':') return { error: `def header must end with ":" — usage: def ${name}(...):` };

  const inner = afterDef.slice(open + 1, close).trim();
  const params: string[] = [];
  if (inner !== '') {
    const tokens = inner.split(/[\s,]+/).filter((t) => t !== '');
    for (const token of tokens) {
      const param = normalizeParam(token);
      if (param === undefined) {
        return { error: `invalid parameter "${token}" in def ${name}(...) — use a plain $name` };
      }
      if (params.includes(param)) {
        return { error: `duplicate parameter "$${param}" in def ${name}(...)` };
      }
      params.push(param);
    }
  }
  return { name, params };
}
