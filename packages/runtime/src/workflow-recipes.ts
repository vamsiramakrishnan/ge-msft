import { z } from 'zod';
import { exactDecimalColumnSql } from '@ge/compute';
import { ReconciliationSpecSchema } from '@ge/contracts';
import { compileAnalysisProgram, type AnalysisProgram } from './analysis-program.js';

const range = z
  .string()
  .trim()
  .min(1)
  .max(1024)
  .refine((value) => !value.includes('```'), 'A range cannot contain a command fence.');
const column = z.number().int().nonnegative().max(16383);
const currency = z.string().regex(/^[A-Z]{3}$/, 'Use a three-letter uppercase currency code.');
const common = { headers: z.boolean().default(true), destination: range.optional() };

export const ReconcileTablesInputsSchema = z
  .object({
    ...common,
    leftRange: range,
    rightRange: range,
    leftKey: column.default(0),
    rightKey: column.default(0),
    leftAmount: column.default(1),
    rightAmount: column.default(1),
    leftCurrency: column.optional(),
    rightCurrency: column.optional(),
    currency: currency.default('USD'),
    tolerance: z
      .string()
      .regex(/^\d{1,12}(\.\d{1,6})?$/)
      .default('0.01'),
  })
  .strict();
export const DuplicateRowsInputsSchema = z
  .object({
    ...common,
    sourceRange: range,
    keyColumn: column.default(0),
    caseSensitive: z.boolean().default(true),
  })
  .strict();
export const SummarizeByGroupInputsSchema = z
  .object({
    ...common,
    sourceRange: range,
    groupColumn: column.default(0),
    amountColumn: column.default(1),
    currency: currency.default('USD'),
    currencyColumn: column.optional(),
  })
  .strict();

export interface WorkflowRecipeField {
  name: string;
  label: string;
  type: 'range' | 'column' | 'decimal' | 'currency' | 'boolean';
  required: boolean;
  default?: string | number | boolean;
  sourceField?: string;
  description: string;
  advanced?: boolean;
}
export interface WorkflowRecipeDefinition {
  id: string;
  version: 1;
  title: string;
  description: string;
  inputSchema: z.ZodType<Record<string, unknown>>;
  fields: readonly WorkflowRecipeField[];
  /** Discovery metadata only. The active session and host remain authoritative. */
  capabilities: readonly ('capture-cells' | 'compute')[];
  writeCapability: 'write-cells';
  result: { binding: 'result'; description: string; emptyMessage: string };
}
const sourceField = (name: string, label: string): WorkflowRecipeField => ({
  name,
  label,
  type: 'range',
  required: true,
  description: 'An explicit source range or named range in the current workbook.',
});
const columnField = (
  name: string,
  label: string,
  source: string,
  fallback?: number,
): WorkflowRecipeField => ({
  name,
  label,
  type: 'column',
  sourceField: source,
  required: fallback !== undefined,
  ...(fallback !== undefined ? { default: fallback } : { advanced: true }),
  description:
    'Zero-based column index within the selected source range. Confirm the captured header before running.',
});
const commonFields: WorkflowRecipeField[] = [
  {
    name: 'headers',
    label: 'First row contains headers',
    type: 'boolean',
    required: true,
    default: true,
    advanced: true,
    description: 'Exclude the first captured row from calculations and use its values as labels.',
  },
  {
    name: 'destination',
    label: 'Write destination',
    type: 'range',
    required: false,
    advanced: true,
    description:
      'Omit to preview. Writing requires a fresh approval and verifies the destination after application.',
  },
];
const currencyField: WorkflowRecipeField = {
  name: 'currency',
  label: 'Currency',
  type: 'currency',
  required: true,
  default: 'USD',
  description:
    'A fixed currency for sources without a selected currency column. A selected currency column takes precedence.',
};

const definitions: WorkflowRecipeDefinition[] = [
  {
    id: 'reconcile-tables',
    version: 1,
    title: 'Reconcile tables',
    description:
      'Match keys and currencies, aggregate exact decimal amounts, and identify variances, missing records and invalid inputs.',
    inputSchema: ReconcileTablesInputsSchema,
    fields: [
      sourceField('leftRange', 'Expected table'),
      sourceField('rightRange', 'Actual table'),
      columnField('leftKey', 'Expected key', 'leftRange', 0),
      columnField('rightKey', 'Actual key', 'rightRange', 0),
      columnField('leftAmount', 'Expected amount', 'leftRange', 1),
      columnField('rightAmount', 'Actual amount', 'rightRange', 1),
      currencyField,
      {
        name: 'tolerance',
        label: 'Allowed variance',
        type: 'decimal',
        required: true,
        default: '0.01',
        description:
          'A nonnegative decimal with at most six fractional digits, applied separately to each key and currency.',
      },
      columnField('leftCurrency', 'Expected currency column', 'leftRange'),
      columnField('rightCurrency', 'Actual currency column', 'rightRange'),
      ...commonFields,
    ],
    capabilities: ['capture-cells', 'compute'],
    writeCapability: 'write-cells',
    result: {
      binding: 'result',
      description:
        'One row per key and currency with totals, variance, status and source row counts. Invalid groups have null totals and variance.',
      emptyMessage: 'Both captured tables contain no data rows. No write is needed.',
    },
  },
  {
    id: 'duplicate-rows',
    version: 1,
    title: 'Find duplicate keys',
    description:
      'List repeated nonblank keys and their counts. Whitespace is trimmed; case sensitivity is an explicit choice.',
    inputSchema: DuplicateRowsInputsSchema,
    fields: [
      sourceField('sourceRange', 'Source table'),
      columnField('keyColumn', 'Key column', 'sourceRange', 0),
      {
        name: 'caseSensitive',
        label: 'Match case',
        type: 'boolean',
        required: true,
        default: true,
        advanced: true,
        description:
          'When disabled, normalize keys to lowercase before grouping. Blank and whitespace-only keys are excluded.',
      },
      ...commonFields,
    ],
    capabilities: ['capture-cells', 'compute'],
    writeCapability: 'write-cells',
    result: {
      binding: 'result',
      description: 'Each repeated key with its occurrence count and number of additional rows.',
      emptyMessage: 'No duplicate nonblank keys found. No write is needed.',
    },
  },
  {
    id: 'summarize-by-group',
    version: 1,
    title: 'Summarize amounts',
    description:
      'Aggregate by group and currency with exact decimal arithmetic. Invalid amounts are counted and prevent a misleading group total.',
    inputSchema: SummarizeByGroupInputsSchema,
    fields: [
      sourceField('sourceRange', 'Source table'),
      columnField('groupColumn', 'Group column', 'sourceRange', 0),
      columnField('amountColumn', 'Amount column', 'sourceRange', 1),
      currencyField,
      columnField('currencyColumn', 'Currency column', 'sourceRange'),
      ...commonFields,
    ],
    capabilities: ['capture-cells', 'compute'],
    writeCapability: 'write-cells',
    result: {
      binding: 'result',
      description:
        'Group, currency, exact total, row count, invalid amount count and validity status. Invalid groups have a null total.',
      emptyMessage: 'The captured table contains no data rows. No write is needed.',
    },
  },
];

/** Return detached metadata; callers cannot mutate the registry through a discovered card. */
export function listWorkflowRecipes(): WorkflowRecipeDefinition[] {
  return definitions.map((definition) => ({
    ...definition,
    fields: definition.fields.map((field) => ({ ...field })),
    capabilities: [...definition.capabilities],
    result: { ...definition.result },
  }));
}
export function getWorkflowRecipe(id: string, version = 1): WorkflowRecipeDefinition {
  const definition = listWorkflowRecipes().find(
    (entry) => entry.id === id && entry.version === version,
  );
  if (!definition)
    throw new Error(
      `Workflow recipe ${id} version ${version} is unavailable. Choose a supported version.`,
    );
  return definition;
}
export function validateWorkflowRecipeInputs(
  id: string,
  inputs: unknown,
  version = 1,
): Record<string, unknown> {
  return getWorkflowRecipe(id, version).inputSchema.parse(inputs);
}

const resultStep = (
  action: Extract<AnalysisProgram['steps'][number], { op: 'bind' }>['action'],
): AnalysisProgram['steps'][number] => ({ op: 'bind', name: 'result', action });
const captureStep = (
  name: string,
  range: string,
  headers: boolean,
): AnalysisProgram['steps'][number] => ({
  op: 'bind',
  name,
  action: { kind: 'capture', range, headers },
});
function complete(steps: AnalysisProgram['steps'], destination?: string): AnalysisProgram {
  if (destination)
    steps.push({ op: 'materialize', id: '$result', destination, whenNonEmpty: true });
  const program: AnalysisProgram = { version: 1, steps, completion: 'verified' };
  compileAnalysisProgram(program);
  return program;
}

/** Deterministic recipe compilation. Inputs are data; no model code or saved authority is accepted. */
export function compileWorkflowRecipe(id: string, raw: unknown, version = 1): AnalysisProgram {
  getWorkflowRecipe(id, version);
  if (id === 'reconcile-tables') {
    const input = ReconcileTablesInputsSchema.parse(raw);
    const spec = ReconciliationSpecSchema.parse({ ...input, left: '$left', right: '$right' });
    return complete(
      [
        captureStep('left', input.leftRange, input.headers),
        captureStep('right', input.rightRange, input.headers),
        resultStep({ kind: 'reconcile', spec }),
      ],
      input.destination,
    );
  }
  if (id === 'duplicate-rows') {
    const input = DuplicateRowsInputsSchema.parse(raw);
    const trimmed = `nullif(trim(cast(c${input.keyColumn} as varchar)), '')`;
    const key = input.caseSensitive ? trimmed : `lower(${trimmed})`;
    return complete(
      [
        captureStep('source', input.sourceRange, input.headers),
        resultStep({
          kind: 'query',
          inputs: ['$source'],
          title: 'Duplicate keys',
          requiredColumns: [{ input: '$source', indices: [input.keyColumn] }],
          sql: `with normalized as (select ${key} as item_key from $source) select item_key, count(*) as occurrences, count(*) - 1 as additional_rows from normalized where item_key is not null group by item_key having count(*) > 1 order by occurrences desc, item_key`,
        }),
      ],
      input.destination,
    );
  }
  const input = SummarizeByGroupInputsSchema.parse(raw);
  const unit =
    input.currencyColumn === undefined
      ? `'${input.currency}'`
      : `upper(trim(cast(c${input.currencyColumn} as varchar)))`;
  // DECIMAL(38,6) accepts 32 integer digits. Reject finer precision and scientific notation;
  // try_cast alone would round them silently. Missing, nonnumeric and overflowing amounts stay invalid.
  const amount = exactDecimalColumnSql(`c${input.amountColumn}`);
  return complete(
    [
      captureStep('source', input.sourceRange, input.headers),
      resultStep({
        kind: 'query',
        inputs: ['$source'],
        title: 'Amounts by group',
        requiredColumns: [
          {
            input: '$source',
            indices: [
              input.groupColumn,
              ...(input.currencyColumn === undefined ? [] : [input.currencyColumn]),
            ],
          },
          { input: '$source', indices: [input.amountColumn], exactDecimal: true },
        ],
        sql: `with normalized as (select nullif(trim(cast(c${input.groupColumn} as varchar)), '') as group_key, ${unit} as currency, ${amount} as amount from $source), grouped as (select group_key, currency, sum(amount) as total, count(*) as records, sum(case when amount is null then 1 else 0 end) as invalid_amounts from normalized group by group_key, currency) select group_key, currency, case when invalid_amounts = 0 and group_key is not null and regexp_full_match(currency, '[A-Z]{3}') then cast(total as varchar) else null end as total, records, invalid_amounts, case when invalid_amounts > 0 or group_key is null or currency is null or not regexp_full_match(currency, '[A-Z]{3}') then 'invalid' else 'valid' end as status from grouped order by status, currency, group_key`,
      }),
    ],
    input.destination,
  );
}

export const WorkflowPresetSchema = z
  .object({
    schemaVersion: z.literal(1),
    recipeId: z.string().min(1).max(64),
    recipeVersion: z.literal(1),
    inputs: z.record(z.unknown()),
  })
  .strict();
export type WorkflowPreset = z.infer<typeof WorkflowPresetSchema>;
/** A preset saves parameters only. It never contains artifacts, approvals, identities or effects. */
export function parseWorkflowPreset(raw: unknown): WorkflowPreset {
  const preset = WorkflowPresetSchema.parse(raw);
  return {
    ...preset,
    inputs: validateWorkflowRecipeInputs(preset.recipeId, preset.inputs, preset.recipeVersion),
  };
}
