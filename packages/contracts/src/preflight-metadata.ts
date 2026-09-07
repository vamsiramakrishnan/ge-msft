import { z } from 'zod';
import { AnalysisActionSchema, ANALYSIS_BINDING_KINDS } from './analysis-actions.js';
import { ContextKindSchema } from './context.js';
import { PlanContextHintSchema } from './command-plan.js';
import { ActuationKindSchema } from './capability.js';
import { approvalClassOf, isReversibleKind } from './plan-graph.js';
import { READ_VERBS, WORKSPACE_VERBS, WRITE_VERB_TO_KIND } from './command-grammar.js';

/** JSON-safe structural guard vocabulary. Unknown Zod features stop generation, never disappear. */
export function describePreflightGuard(schema: z.ZodTypeAny): Record<string, unknown> {
  if (schema instanceof z.ZodOptional)
    return { ...describePreflightGuard(schema.unwrap()), optional: true };
  if (schema instanceof z.ZodBoolean) return { type: 'boolean' };
  if (schema instanceof z.ZodString || schema instanceof z.ZodNumber) {
    if (schema._def.coerce) throw new Error('Unsupported preflight guard coercion');
    const number = schema instanceof z.ZodNumber;
    const allowed = number ? ['min', 'max', 'int', 'finite'] : ['min', 'max'];
    for (const check of schema._def.checks)
      if (!allowed.includes(check.kind))
        throw new Error(`Unsupported preflight guard check ${check.kind}`);
    return {
      type: number ? 'number' : 'string',
      checks: schema._def.checks.map((check) =>
        check.kind === 'min' || check.kind === 'max'
          ? {
              kind: check.kind,
              value: check.value,
              ...('inclusive' in check ? { inclusive: check.inclusive } : {}),
            }
          : { kind: check.kind },
      ),
    };
  }
  if (schema instanceof z.ZodArray) {
    if (schema._def.exactLength) throw new Error('Unsupported preflight guard exact array length');
    return {
      type: 'array',
      items: describePreflightGuard(schema.element),
      min: schema._def.minLength?.value,
      max: schema._def.maxLength?.value,
    };
  }
  if (schema instanceof z.ZodObject) {
    if (!(schema._def.catchall instanceof z.ZodNever))
      throw new Error('Unsupported preflight guard catchall');
    return {
      type: 'object',
      strict: schema._def.unknownKeys === 'strict',
      properties: Object.fromEntries(
        Object.entries(schema.shape as Record<string, z.ZodTypeAny>).map(([key, value]) => [
          key,
          describePreflightGuard(value),
        ]),
      ),
    };
  }
  throw new Error(`Unsupported preflight guard schema ${schema._def.typeName}`);
}

/** Advisory compiler metadata derived from the same contracts as runtime dispatch and approval. */
export function buildPreflightMetadata() {
  const query = AnalysisActionSchema.options.find((action) => action.shape.kind.value === 'query')!;
  const materialize = AnalysisActionSchema.options.find(
    (action) => action.shape.kind.value === 'materialize',
  )!;
  const fields = (schema: z.AnyZodObject, names: string[]) =>
    Object.fromEntries(
      names.map((name) => [name, describePreflightGuard(schema.shape[name] as z.ZodTypeAny)]),
    );
  return {
    formatVersion: 1 as const,
    contextKinds: [...ContextKindSchema.options],
    contextHints: [...PlanContextHintSchema.options],
    analysisBindingKinds: [...ANALYSIS_BINDING_KINDS],
    readPhaseVerbs: [
      ...READ_VERBS,
      ...WORKSPACE_VERBS.filter((verb) => ['workspace', 'save', 'cat', 'grep'].includes(verb)),
    ],
    approvalByKind: Object.fromEntries(
      ActuationKindSchema.options.map((kind) => [
        kind,
        { approvalClass: approvalClassOf(kind), reversible: isReversibleKind(kind) },
      ]),
    ),
    approvalByVerb: {
      ...Object.fromEntries(
        Object.entries(WRITE_VERB_TO_KIND).map(([verb, kind]) => [
          verb,
          { approvalClass: approvalClassOf(kind), reversible: isReversibleKind(kind) },
        ]),
      ),
      share: { approvalClass: 'estate' as const, reversible: false },
    },
    analysisGuards: {
      query: fields(query, ['requiredColumns']),
      materialize: fields(materialize, ['whenNonEmpty']),
    },
  };
}
