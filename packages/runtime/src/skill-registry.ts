import {
  parseProgramBlock,
  shadowsBuiltin,
  type ParsedSkillCall,
  type ParsedSkillDef,
  type ProgramEntry,
} from '@ge/contracts';

/**
 * ADR-0005 Phase 3 — the in-session skill registry + macro expansion.
 *
 * A skill is a named, parameterized composition the model defines once (`def name(p…): … end`) and
 * calls (`name arg…`). This module is the runtime side: it stores definitions in an in-session Map
 * and, on a call, binds positional args → params, textually substitutes `$param` tokens in the
 * body lines, and returns the expanded **lines** — which the caller re-parses through the EXISTING
 * Phase-2 plan machinery (type-check → dry-run → approvePlan → gated execute). A skill call is thus
 * *just a plan*: it introduces no new gate, no approval bypass, no effect that skips dry-run.
 *
 * Durable persistence (skills surviving a session, in host metadata) is a deliberate follow-up;
 * this wave keeps the registry in-session (the Map lives for the life of the {@link SkillRegistry}).
 *
 * SCOPE (this wave): parameterized macros — substitution only, NO `for`/`each` iteration.
 *
 * SECURITY — substitution is textual into ALREADY-PARSED-AND-RE-TYPE-CHECKED lines, with two hard
 * rules so a malicious arg can never smuggle an effect past the gate:
 *   • an argument may NOT contain a newline (it can therefore never inject a NEW command line —
 *     `slide "x"\npost "evil"` is rejected), and
 *   • the expansion is BOUNDED (`maxBodyLines`) so a skill body can't fan out unboundedly; the
 *     per-plan effect cap in the loop is the second bound on the resulting effects.
 * The expanded lines are re-parsed and every effect among them still flows through dry-run +
 * approvePlan + the actuation gate — substitution changes WHAT a line says, never WHETHER it gates.
 */

/** A stored skill: its declared params (no `$`) and its verbatim body lines. */
export interface RegisteredSkill {
  name: string;
  params: string[];
  body: string[];
}

/** The outcome of registering a `def` — a confirmation (no execution) or a corrective error. */
export type RegisterResult =
  | { ok: true; name: string; params: string[]; bodyLines: number; redefined: boolean }
  | { ok: false; error: string };

/** The outcome of expanding a call — the substituted body lines, or a corrective error. */
export type ExpandResult = { ok: true; lines: string[] } | { ok: false; error: string };

/** Defensive cap on a single skill body's size (also bounds the expanded-line count). */
const DEFAULT_MAX_BODY_LINES = 64;

export interface SkillRegistryOptions {
  /** Max body lines a skill may have / a call may expand to (default {@link DEFAULT_MAX_BODY_LINES}). */
  maxBodyLines?: number;
}

export class SkillRegistry {
  private readonly skills = new Map<string, RegisteredSkill>();
  private readonly maxBodyLines: number;

  constructor(opts: SkillRegistryOptions = {}) {
    this.maxBodyLines = opts.maxBodyLines ?? DEFAULT_MAX_BODY_LINES;
  }

  /** The currently-registered skill names — fed to `parseProgramBlock` so calls parse as calls. */
  names(): ReadonlySet<string> {
    return new Set(this.skills.keys());
  }

  has(name: string): boolean {
    return this.skills.has(name);
  }

  get(name: string): RegisteredSkill | undefined {
    return this.skills.get(name);
  }

  /**
   * Register a `def`. Rejects (defense-in-depth — the parser already enforces these) a name that
   * shadows a built-in verb, a duplicate parameter, or an over-long body. A `$param` referenced in
   * the body that is not declared is rejected here (a clear corrective at define-time beats an
   * unbound-token surprise at call-time). Registering does NOT execute the body — it returns a
   * confirmation. Re-defining an existing name replaces it (and is flagged `redefined`).
   */
  register(def: ParsedSkillDef): RegisterResult {
    if (shadowsBuiltin(def.name)) {
      return {
        ok: false,
        error: `skill "${def.name}" shadows a built-in command — choose another name`,
      };
    }
    if (new Set(def.params).size !== def.params.length) {
      return { ok: false, error: `skill "${def.name}" has duplicate parameters` };
    }
    if (def.body.length === 0) {
      return { ok: false, error: `skill "${def.name}" has an empty body` };
    }
    if (def.body.length > this.maxBodyLines) {
      return {
        ok: false,
        error: `skill "${def.name}" body is too long (${def.body.length} > ${this.maxBodyLines} lines)`,
      };
    }
    // A `$token` in the body must be either a declared param OR a name the body itself binds with
    // `let $x = …` (an intermediate value, ADR-0005 Phase 1) — anything else is a typo/unbound
    // reference, caught here as a clear define-time corrective rather than at call-time.
    const known = new Set([...def.params, ...letBoundNames(def.body)]);
    const undeclared = referencedParams(def.body).filter((p) => !known.has(p));
    if (undeclared.length > 0) {
      return {
        ok: false,
        error: `skill "${def.name}" references undeclared parameter(s): ${undeclared
          .map((p) => `$${p}`)
          .join(', ')}`,
      };
    }

    const redefined = this.skills.has(def.name);
    this.skills.set(def.name, { name: def.name, params: def.params, body: [...def.body] });
    return { ok: true, name: def.name, params: def.params, bodyLines: def.body.length, redefined };
  }

  /**
   * Expand a call into its substituted body lines, ready to re-parse + run as a plan. Checks:
   *   • the skill is registered (undefined-name → corrective);
   *   • arity matches exactly (too few / too many args → corrective);
   *   • no argument contains a newline (so a substituted arg can NEVER inject a new command line).
   * Substitutes each declared `$param` token in every body line with its bound arg, then returns
   * the expanded lines. The expansion is bounded by the body size; the caller's per-plan effect cap
   * bounds the resulting effects. Never throws — a bad expansion is a corrective `{ ok: false }`.
   */
  expand(call: ParsedSkillCall): ExpandResult {
    const skill = this.skills.get(call.name);
    if (!skill) {
      return { ok: false, error: `unknown skill "${call.name}" — define it first with def` };
    }
    if (call.args.length !== skill.params.length) {
      return {
        ok: false,
        error: `skill "${call.name}" expects ${skill.params.length} argument(s) (${skill.params
          .map((p) => `$${p}`)
          .join(', ')}), got ${call.args.length}`,
      };
    }
    for (let i = 0; i < call.args.length; i++) {
      const arg = call.args[i]!;
      if (/[\r\n]/.test(arg)) {
        return {
          ok: false,
          error: `argument ${i + 1} to "${call.name}" contains a newline — arguments must be single-line`,
        };
      }
      // A literal triple-backtick would close the synthetic ```cmd fence early on re-parse
      // (reparseExpandedLines), silently truncating the rest of the expansion — reject it.
      if (arg.includes('```')) {
        return {
          ok: false,
          error: `argument ${i + 1} to "${call.name}" contains a code fence — not allowed in an argument`,
        };
      }
    }
    if (skill.body.length > this.maxBodyLines) {
      return {
        ok: false,
        error: `skill "${call.name}" body exceeds the ${this.maxBodyLines}-line cap`,
      };
    }

    const bindings = new Map<string, string>();
    for (let i = 0; i < skill.params.length; i++) bindings.set(skill.params[i]!, call.args[i]!);

    const lines = skill.body.map((line) => substituteParams(line, bindings));
    return { ok: true, lines };
  }
}

/** A `$name` token reference regex — the param identifier rule mirrors the grammar's. */
const PARAM_TOKEN_RE = /\$([A-Za-z_][A-Za-z0-9_]*)/g;

/** Every distinct `$name` referenced across a set of body lines (for the undeclared-param check). */
function referencedParams(body: string[]): string[] {
  const found = new Set<string>();
  for (const line of body) {
    for (const m of line.matchAll(PARAM_TOKEN_RE)) found.add(m[1]!);
  }
  return [...found];
}

/** Names a body binds locally via `let $x = …` (ADR-0005 Phase 1 intermediate values). */
const LET_BIND_RE = /^let\s+\$([A-Za-z_][A-Za-z0-9_]*)\s*=/i;
function letBoundNames(body: string[]): string[] {
  const names: string[] = [];
  for (const line of body) {
    const m = LET_BIND_RE.exec(line.trim());
    if (m) names.push(m[1]!);
  }
  return names;
}

/**
 * Substitute every declared `$param` token in a line with its bound argument. Only EXACT declared
 * parameter names are replaced — a `$other` that is not a parameter is left untouched (so a literal
 * `$` in a value is not mangled). Replacement is by name match on the whole identifier (the regex's
 * `[A-Za-z0-9_]*` greedily consumes the identifier, so `$ab` never matches a `$a` binding).
 *
 * The bound value is inserted VERBATIM (already guaranteed newline-free by {@link SkillRegistry.expand}),
 * so the line stays a single line and re-parses as exactly one command — an arg cannot inject a new
 * effect line.
 */
function substituteParams(line: string, bindings: ReadonlyMap<string, string>): string {
  return line.replace(PARAM_TOKEN_RE, (whole, name: string) => {
    const bound = bindings.get(name);
    return bound !== undefined ? bound : whole;
  });
}

/**
 * Re-parse a skill's expanded body lines back into program entries (commands / expressions),
 * scoped to the registry's other skills so a body may call another already-defined skill. A `def`
 * inside an expanded body is NOT allowed (the parser's nested-def guard already rejects it). The
 * caller runs the resulting entries through the existing Phase-2 plan machinery.
 *
 * Implemented by re-framing the lines as a synthetic ```cmd fence and delegating to the single
 * source of truth, {@link parseProgramBlock} — so expanded lines parse EXACTLY as if the model had
 * emitted them, with zero parser duplication.
 */
export function reparseExpandedLines(
  lines: string[],
  knownSkills: ReadonlySet<string>,
): ProgramEntry[] {
  const fenced = '```cmd\n' + lines.join('\n') + '\n```';
  const { entries } = parseProgramBlock(fenced, knownSkills);
  return entries;
}
