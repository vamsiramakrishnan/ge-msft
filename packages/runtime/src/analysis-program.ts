import { z } from 'zod';
import { validateQuery } from '@ge/compute';
import { AnalysisActionSchema, type AnalysisAction } from './analysis-workspace.js';

const NAME = /^[A-Za-z_][A-Za-z0-9_]*$/;
const PRODUCES_ARTIFACT = new Set(['capture', 'query', 'reconcile', 'filter', 'inspect']);
export const AnalysisProgramSchema = z
  .object({
    version: z.literal(1),
    steps: z
      .array(
        z.discriminatedUnion('op', [
          z.object({
            op: z.literal('bind'),
            name: z.string().max(64).regex(NAME),
            action: AnalysisActionSchema,
          }),
          z.object({
            op: z.literal('materialize'),
            id: z.string().min(1),
            destination: z.string().min(1),
            whenNonEmpty: z.boolean().optional(),
          }),
        ]),
      )
      .min(1)
      .max(31),
    completion: z.literal('verified').default('verified'),
  })
  .strict();
export type AnalysisProgram = z.input<typeof AnalysisProgramSchema>;

/** Task-local, typed references. A name never substitutes into targets, literals or arbitrary JSON. */
export class AnalysisBindings {
  private readonly values = new Map<string, string>();
  clear(): void {
    this.values.clear();
  }
  has(name: string): boolean {
    return this.values.has(name);
  }
  entries(): readonly (readonly [string, string])[] {
    return [...this.values.entries()].map(([name, id]) => [name, id] as const);
  }
  bind(name: string, id: string): void {
    if (!NAME.test(name) || name.length > 64) throw new Error('Invalid analysis binding name.');
    if (!/^a_[a-f0-9]{24}$/.test(id)) throw new Error('Only artifact handles may be bound.');
    if (this.values.has(name))
      throw new Error(`Binding $${name} already exists. Choose a new name.`);
    this.values.set(name, id);
  }
  resolve(raw: unknown): AnalysisAction {
    const action = AnalysisActionSchema.parse(raw);
    const ref = (value: string): string => {
      if (!value.startsWith('$')) return value;
      const found = this.values.get(value.slice(1));
      if (!found) throw new Error(`Unknown artifact binding ${value}. Bind it before use.`);
      return found;
    };
    if (action.kind === 'reconcile')
      return {
        ...action,
        spec: { ...action.spec, left: ref(action.spec.left), right: ref(action.spec.right) },
      };
    if (action.kind === 'query') {
      const inputs = action.inputs.map(ref);
      return {
        ...action,
        inputs,
        sql: resolveSqlBindings(action.sql, this.values, new Set(inputs)),
        ...(action.requiredColumns
          ? {
              requiredColumns: action.requiredColumns.map((entry) => ({
                ...entry,
                input: ref(entry.input),
              })),
            }
          : {}),
      };
    }
    if (
      action.kind === 'inspect' ||
      action.kind === 'filter' ||
      action.kind === 'materialize' ||
      action.kind === 'remove'
    )
      return { ...action, id: ref(action.id) };
    return action;
  }
}

/** SQL parameters here denote admitted tables only. Quoted strings are copied byte for byte. */
function resolveSqlBindings(
  sql: string,
  bindings: ReadonlyMap<string, string>,
  inputs: ReadonlySet<string>,
): string {
  let out = '';
  let i = 0;
  while (i < sql.length) {
    if (sql[i] === "'") {
      const start = i++;
      while (i < sql.length) {
        if (sql[i++] === "'") {
          if (sql[i] === "'") {
            i++;
            continue;
          }
          break;
        }
      }
      out += sql.slice(start, i);
    } else if (sql[i] === '$') {
      const match = /^\$([A-Za-z_][A-Za-z0-9_]*)/.exec(sql.slice(i));
      if (!match || (i > 0 && /[A-Za-z0-9_]/.test(sql[i - 1]!)))
        throw new Error('Invalid SQL artifact binding.');
      const id = bindings.get(match[1]!);
      if (!id || !inputs.has(id))
        throw new Error(`SQL binding ${match[0]} must be present in query inputs.`);
      out += id;
      i += match[0].length;
    } else out += sql[i++];
  }
  return validateQuery(out);
}

/** Compile SDK input to the same CLI program used by model-authored and pasted programs. */
export function compileAnalysisProgram(raw: AnalysisProgram): string {
  const program = AnalysisProgramSchema.parse(raw);
  const names = new Set<string>();
  const lines: string[] = [];
  for (const step of program.steps) {
    const action: AnalysisAction =
      step.op === 'bind'
        ? step.action
        : {
            kind: 'materialize',
            id: step.id,
            destination: step.destination,
            ...(step.whenNonEmpty !== undefined ? { whenNonEmpty: step.whenNonEmpty } : {}),
          };
    if (step.op === 'bind' && !PRODUCES_ARTIFACT.has(action.kind))
      throw new Error('Only artifact-producing actions can be bound.');
    if (step.op === 'bind' && names.has(step.name))
      throw new Error(`Duplicate binding $${step.name}.`);
    const refs =
      action.kind === 'query'
        ? action.inputs
        : action.kind === 'reconcile'
          ? [action.spec.left, action.spec.right]
          : 'id' in action
            ? [action.id]
            : [];
    for (const ref of refs)
      if (ref.startsWith('$') && !names.has(ref.slice(1)))
        throw new Error(`Unbound artifact ${ref}.`);
    const encoded = JSON.stringify(action);
    if (encoded.includes('```'))
      throw new Error('Program values cannot contain a command fence. Use an artifact reference.');
    lines.push(
      step.op === 'bind' ? `let $${step.name} = analyze ${encoded}` : `analyze ${encoded}`,
    );
    if (step.op === 'bind') names.add(step.name);
  }
  lines.push('finish when=verified');
  return lines.join('\n');
}

export interface AnalysisProgramPlan {
  steps: Array<{
    index: number;
    op: 'bind' | 'materialize';
    effect: 'read' | 'derive' | 'write';
    dependsOn: number[];
    binding?: string;
  }>;
  layers: number[][];
  independentCaptureGroups: number[][];
  execution: 'serial';
  reason: string;
}

/** Exposes data dependencies without pretending that DocBridge offers atomic batch capture. */
export function inspectAnalysisProgram(raw: AnalysisProgram): AnalysisProgramPlan {
  compileAnalysisProgram(raw);
  const program = AnalysisProgramSchema.parse(raw);
  const bindings = new Map<string, number>();
  const steps: AnalysisProgramPlan['steps'] = [];
  const levels: number[] = [];
  let lastWrite: number | undefined;
  for (const [index, step] of program.steps.entries()) {
    const action = step.op === 'bind' ? step.action : undefined;
    const refs =
      step.op === 'materialize'
        ? [step.id]
        : action?.kind === 'query'
          ? action.inputs
          : action?.kind === 'reconcile'
            ? [action.spec.left, action.spec.right]
            : action && 'id' in action
              ? [action.id]
              : [];
    const dependencies = new Set<number>();
    for (const ref of refs) if (ref.startsWith('$')) dependencies.add(bindings.get(ref.slice(1))!);
    if (lastWrite !== undefined) dependencies.add(lastWrite);
    if (step.op === 'materialize') {
      // Preserve effect order against all earlier reads, including currently independent captures.
      for (let previous = 0; previous < index; previous++) dependencies.add(previous);
      lastWrite = index;
    }
    const dependsOn = [...dependencies].sort((a, b) => a - b);
    steps.push({
      index,
      op: step.op,
      effect: step.op === 'materialize' ? 'write' : action?.kind === 'capture' ? 'read' : 'derive',
      dependsOn,
      ...(step.op === 'bind' ? { binding: step.name } : {}),
    });
    levels.push(
      dependsOn.length ? 1 + Math.max(...dependsOn.map((dependency) => levels[dependency]!)) : 0,
    );
    if (step.op === 'bind') bindings.set(step.name, index);
  }
  const layers: number[][] = [];
  levels.forEach((level, index) => (layers[level] ??= []).push(index));
  return {
    steps,
    layers,
    independentCaptureGroups: layers
      .map((layer) => layer.filter((index) => steps[index]!.effect === 'read'))
      .filter((group) => group.length > 1),
    execution: 'serial',
    reason:
      'The current host bridge exposes individual versioned captures. Execution preserves serial host access and rechecks sources before computation and writes.',
  };
}
