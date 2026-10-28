import {
  isProgramAnalysisBinding,
  isProgramCommand,
  isProgramEachBlock,
  isProgramExpr,
  isProgramSkillCall,
  isProgramSkillDef,
  type ParsedCommand,
  type ParsedExpr,
  type PipelineExpr,
  type ProgramEntry,
} from '@ge/contracts';
import { EFFECT_VERBS } from '@ge/contracts';
import { TRANSFORMS } from './compose.js';
import { reparseExpandedLines, SkillRegistry } from './skill-registry.js';

/**
 * B2 — the STATIC type pass over a parsed program, run before dry-run (ADR-0005 Phase 2 type-check
 * → dry-run → preview → approval). It infers a value type for every binding and pipeline without
 * touching the host, and rejects the programs the evaluator would only refuse mid-turn — or, worse,
 * mid-plan after earlier lines already read the document:
 *
 *   • an unknown transform (`| sumx amount`), with a did-you-mean;
 *   • a transform fed the wrong kind of value (`… | sum amount | head 3` — a number has no rows);
 *   • a `$name` used before it is bound (`set A1 = ($ghost | sum x)`), including inside an effect's
 *     embedded expression and an `each … in $table` header;
 *   • a column law violation where the columns are statically known (a table literal, or a table
 *     derived from one through select/rename/derive/group): `select`, `sort`, `sum`/`avg`/`min`/
 *     `max`, `filter`, `group`, `distinct`, `rename` and `join … on` must name columns the table has.
 *
 * The pass is CONSERVATIVE: a host read's columns are unknown until the read runs, so column laws
 * are only applied to tables whose columns the program itself determines. Types are
 * `Table | Number | Text | Artifact | Unknown`; the Python preflight (`surface_cli`) infers the same
 * `Table`/`Number`/`Text` for bindings — keep the two in step.
 */
export type StaticType = 'Table' | 'Number' | 'Text' | 'Artifact' | 'Unknown';

export interface StaticBinding {
  type: StaticType;
  /** The table's column names when the program determines them; absent when a host read decides. */
  columns?: readonly string[];
}

export interface StaticDiagnostic {
  /** Index of the offending entry in the program. */
  index: number;
  message: string;
}

/** Transforms that reduce a table to a number. */
const AGGREGATES = new Set(['sum', 'avg', 'min', 'max', 'count']);
/** Transforms that keep a text value a text value (and a table a table). */
const TEXT_OK = new Set(['sed']);

function closest(name: string, candidates: Iterable<string>): string | undefined {
  let best: { name: string; distance: number } | undefined;
  for (const candidate of candidates) {
    const distance = levenshtein(name.toLowerCase(), candidate.toLowerCase());
    if (!best || distance < best.distance) best = { name: candidate, distance };
  }
  return best && best.distance <= Math.max(2, Math.floor(name.length / 3)) ? best.name : undefined;
}

function levenshtein(a: string, b: string): number {
  const dp = Array.from({ length: a.length + 1 }, (_, i) => [
    i,
    ...Array<number>(b.length).fill(0),
  ]);
  for (let j = 1; j <= b.length; j++) dp[0]![j] = j;
  for (let i = 1; i <= a.length; i++)
    for (let j = 1; j <= b.length; j++)
      dp[i]![j] = Math.min(
        dp[i - 1]![j]! + 1,
        dp[i]![j - 1]! + 1,
        dp[i - 1]![j - 1]! + (a[i - 1] === b[j - 1] ? 0 : 1),
      );
  return dp[a.length]![b.length]!;
}

const splitList = (text: string): string[] =>
  text
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);

/** The columns a transform's args name, by transform (only the shapes the evaluator accepts). */
function referencedColumns(name: string, args: string): string[] {
  const first = args.trim().split(/\s+/)[0] ?? '';
  switch (name) {
    case 'select':
    case 'distinct':
      return splitList(args);
    case 'sort':
    case 'sum':
    case 'avg':
    case 'min':
    case 'max':
      return first ? [first] : [];
    case 'filter':
      return [...args.matchAll(/(?:^|[\s(])([A-Za-z_][\w]*)\s*(?:>=|<=|!=|contains|=|>|<)/g)].map(
        (m) => m[1]!,
      );
    case 'rename':
      return splitList(args).map((pair) => pair.split('=')[0]!.trim());
    case 'group': {
      const m = /^(.*?)\s+agg\s+/i.exec(args.trim());
      return m ? splitList(m[1]!) : [];
    }
    case 'join': {
      const m = /\bon\s+([A-Za-z_][\w]*)/i.exec(args);
      return m ? [m[1]!] : [];
    }
    default:
      return [];
  }
}

/** The columns a table has AFTER a transform, when they can be known from the columns before. */
function resultColumns(
  name: string,
  args: string,
  before: readonly string[] | undefined,
): readonly string[] | undefined {
  if (!before) {
    if (name === 'select') return splitList(args);
    return undefined;
  }
  switch (name) {
    case 'select':
      return splitList(args);
    case 'rename': {
      const map = new Map(
        splitList(args).map((pair) => {
          const [from, to] = pair.split('=').map((s) => s.trim());
          return [from ?? '', to ?? from ?? ''];
        }),
      );
      return before.map((c) => map.get(c) ?? c);
    }
    case 'derive': {
      const m = /^([A-Za-z_][\w]*)\s*=/.exec(args.trim());
      return m ? [...before, m[1]!] : undefined;
    }
    case 'group': {
      const m = /^(.*?)\s+agg\s+(.*)$/i.exec(args.trim());
      if (!m) return undefined;
      const names = [...m[2]!.matchAll(/\bas\s+([A-Za-z_][\w]*)/gi)].map((x) => x[1]!);
      return [...splitList(m[1]!), ...names];
    }
    case 'join':
      return undefined;
    default:
      return before;
  }
}

/** Infer a pipeline's static type, reporting the first type error it would hit. */
export function inferPipeline(
  expr: PipelineExpr,
  env: ReadonlyMap<string, StaticBinding>,
): { binding: StaticBinding; error?: string } {
  let current: StaticBinding;
  const src = expr.source;
  switch (src.src) {
    case 'number':
      current = { type: 'Number' };
      break;
    case 'text':
      current = { type: 'Text' };
      break;
    case 'table':
      current = { type: 'Table', columns: src.columns };
      break;
    case 'var': {
      const bound = env.get(src.name);
      if (!bound)
        return {
          binding: { type: 'Unknown' },
          error: `$${src.name} is an unbound variable (used before it is bound)`,
        };
      if (bound.type === 'Artifact')
        return {
          binding: { type: 'Unknown' },
          error: `$${src.name} is an analysis artifact; use analyze to read it`,
        };
      current = bound;
      break;
    }
    default:
      // A host read: a table whose columns only the read can tell.
      current = { type: 'Table' };
  }
  for (const stage of expr.stages) {
    const name = stage.name.toLowerCase();
    // Piping INTO an effect (`read A1 | set B1 …`) is the evaluator's Phase-1 corrective ("effects
    // can't be composed"); leave it to dry-run so the model gets that exact wording.
    if (EFFECT_VERBS.has(name as never)) return { binding: { type: 'Unknown' } };
    if (!(name in TRANSFORMS)) {
      const hint = closest(name, Object.keys(TRANSFORMS));
      return {
        binding: { type: 'Unknown' },
        error: `unknown transform "${stage.name}"${hint ? ` — did you mean "${hint}"?` : ''}`,
      };
    }
    if (current.type === 'Number' || (current.type === 'Text' && !TEXT_OK.has(name)))
      return {
        binding: { type: 'Unknown' },
        error: `"${stage.name}" needs a table, but the value before it is a ${current.type.toLowerCase()}`,
      };
    if (current.type === 'Table' && current.columns) {
      const missing = referencedColumns(name, stage.args).filter(
        (c) => !current.columns!.includes(c),
      );
      if (missing.length)
        return {
          binding: { type: 'Unknown' },
          error: `"${stage.name}" names column${missing.length > 1 ? 's' : ''} ${missing
            .map((c) => `"${c}"`)
            .join(', ')} the table does not have (columns: ${current.columns.join(', ')})`,
        };
    }
    if (name === 'join') {
      const other = /^\$([A-Za-z_][\w]*)/.exec(stage.args.trim())?.[1];
      const bound = other ? env.get(other) : undefined;
      if (other && !bound)
        return {
          binding: { type: 'Unknown' },
          error: `$${other} is an unbound variable (used before it is bound)`,
        };
      if (bound && bound.type !== 'Table' && bound.type !== 'Unknown')
        return {
          binding: { type: 'Unknown' },
          error: `join needs a bound table, but $${other} is a ${bound.type.toLowerCase()}`,
        };
    }
    if (AGGREGATES.has(name)) current = { type: 'Number' };
    else if (current.type === 'Text') current = { type: 'Text' };
    else current = { type: 'Table', columns: resultColumns(name, stage.args, current.columns) };
  }
  return { binding: current };
}

/** The expression embedded in an effect command (`set A1 = (…)`, `slide "t" (…)`, …), if any. */
function embeddedExpr(command: ParsedCommand): ParsedExpr | undefined {
  const c = command as Partial<
    Record<'valueExpr' | 'textExpr' | 'bulletsExpr' | 'bodyExpr', ParsedExpr>
  >;
  return c.valueExpr ?? c.textExpr ?? c.bulletsExpr ?? c.bodyExpr;
}

/**
 * Type-check a program against the bindings already in scope. Returns every diagnostic (one per
 * offending entry, first error in that entry) and the environment the program would leave behind.
 * Pure: no host, no evaluation.
 */
export function typecheckProgram(
  entries: readonly ProgramEntry[],
  scope: ReadonlyMap<string, StaticBinding>,
  options: { skills?: SkillRegistry; maxSkillDepth?: number; maxCommands?: number } = {},
): { diagnostics: StaticDiagnostic[]; env: Map<string, StaticBinding> } {
  const env = new Map(scope);
  const diagnostics: StaticDiagnostic[] = [];
  // A definition affects later calls, but this preflight must not mutate the live registry. Seed
  // a private copy with the recipes already loaded from the document or registered in-session.
  const skills = new SkillRegistry();
  for (const name of options.skills?.names() ?? []) {
    skills.register({ kind: 'skill-def', ...options.skills!.get(name)! });
  }
  let remaining = options.maxCommands ?? 4096;
  const maxDepth = options.maxSkillDepth ?? 8;
  const visit = (entry: ProgramEntry, index: number, depth: number): void => {
    if (remaining <= 0) return;
    remaining -= 1;
    if (isProgramSkillDef(entry)) {
      // Registration and expansion errors are reported by the execution path itself.
      skills.register(entry);
      return;
    }
    if (isProgramSkillCall(entry)) {
      if (depth >= maxDepth) return;
      const expanded = skills.expand(entry);
      if (!expanded.ok) return;
      for (const child of reparseExpandedLines(expanded.lines, skills.names()))
        visit(child, index, depth + 1);
      return;
    }
    if (isProgramExpr(entry)) {
      const pipeline = entry.kind === 'let' ? entry.pipeline : entry;
      const { binding, error } = inferPipeline(pipeline, env);
      if (error) diagnostics.push({ index, message: error });
      if (entry.kind === 'let') env.set(entry.name, error ? { type: 'Unknown' } : binding);
      return;
    }
    if (isProgramAnalysisBinding(entry)) {
      env.set(entry.name, { type: 'Artifact' });
      return;
    }
    if (isProgramEachBlock(entry)) {
      const table = env.get(entry.tableVar);
      if (!table)
        diagnostics.push({
          index,
          message: `$${entry.tableVar} is an unbound variable (used before it is bound)`,
        });
      else if (table.type !== 'Table' && table.type !== 'Unknown')
        diagnostics.push({
          index,
          message: `each needs a table, but $${entry.tableVar} is a ${table.type.toLowerCase()}`,
        });
      else if (table.columns) {
        const refs = entry.body.flatMap((line) =>
          [...line.matchAll(new RegExp(`\\$${entry.rowVar}\\.([A-Za-z_][\\w]*)`, 'g'))].map(
            (m) => m[1]!,
          ),
        );
        const missing = [...new Set(refs.filter((c) => !table.columns!.includes(c)))];
        if (missing.length)
          diagnostics.push({
            index,
            message: `each body names column${missing.length > 1 ? 's' : ''} ${missing
              .map((c) => `"${c}"`)
              .join(
                ', ',
              )} that $${entry.tableVar} does not have (columns: ${table.columns.join(', ')})`,
          });
      }
      return;
    }
    if (isProgramCommand(entry)) {
      const expr = embeddedExpr(entry);
      if (!expr) return;
      const pipeline = expr.kind === 'let' ? expr.pipeline : expr;
      const { error } = inferPipeline(pipeline, env);
      if (error) diagnostics.push({ index, message: error });
    }
  };
  entries.forEach((entry, index) => visit(entry, index, 0));
  return { diagnostics, env };
}

/** The static view of the live evaluator bindings (what a running loop already holds). */
export function staticScope(
  values: ReadonlyMap<string, { kind: 'table' | 'number' | 'text'; columns?: string[] }>,
  artifacts: Iterable<string> = [],
): Map<string, StaticBinding> {
  const scope = new Map<string, StaticBinding>();
  for (const [name, value] of values)
    scope.set(
      name,
      value.kind === 'table'
        ? { type: 'Table', columns: value.columns }
        : { type: value.kind === 'number' ? 'Number' : 'Text' },
    );
  for (const name of artifacts) scope.set(name, { type: 'Artifact' });
  return scope;
}
