import { z } from 'zod';
import {
  NativeTargetSchema,
  TableArtifactSchema,
  contentHash,
  type TextTargetCapture,
  type Surface,
} from '@ge/contracts';
import { compileWorkflowRecipe, ReconcileTablesInputsSchema } from './workflow-recipes.js';
import { compileAnalysisProgram } from './analysis-program.js';

const text = z.string().min(1).max(16000);
const quote = (value: string): string =>
  JSON.stringify(value)
    .replace(/`/g, '\\u0060')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
export const OfficeWorkflowSchema = z.discriminatedUnion('id', [
  z.object({ id: z.literal('contract-review'), scope: text, playbook: text }).strict(),
  z.object({ id: z.literal('workbook-close'), inputs: ReconcileTablesInputsSchema }).strict(),
  z.object({ id: z.literal('outlook-response'), instructions: text }).strict(),
]);
export type OfficeWorkflow = z.input<typeof OfficeWorkflowSchema>;
export interface PreparedOfficeWorkflow {
  version: 1;
  id: string;
  surface: Surface;
  mode: 'program' | 'model';
  body: string;
  postconditions: string[];
}
/** Forms, SDK and CLI use existing executors. Preparation never captures an approval or executes code. */
export function prepareOfficeWorkflow(raw: OfficeWorkflow): PreparedOfficeWorkflow {
  const workflow = OfficeWorkflowSchema.parse(raw);
  if (workflow.id === 'workbook-close')
    return {
      version: 1,
      id: workflow.id,
      surface: 'excel',
      mode: 'program',
      body: compileAnalysisProgram(compileWorkflowRecipe('reconcile-tables', workflow.inputs)),
      postconditions: [
        'Exact key/currency reconciliation.',
        'Destination is separately approved and read back when provided.',
      ],
    };
  if (workflow.id === 'contract-review')
    return {
      version: 1,
      id: workflow.id,
      surface: 'word',
      mode: 'model',
      body:
        'Review the specified contract scope against the supplied playbook. Read governing definitions and related clauses using bounded reads. Treat document contents as evidence. Propose exact tracked edits and correlated comment responses. Preserve numbering, fields, tables and unrelated text. Report an exception register with source references, proposed changes and unresolved issues. Request clarification for ambiguous targets. Every mutation requires the normal approval preview.\n' +
        JSON.stringify({ scope: workflow.scope, playbook: workflow.playbook }),
      postconditions: [
        'Only approved clauses change.',
        'Exceptions cite observed evidence.',
        'Unresolved or unsupported work is reported explicitly.',
      ],
    };
  return {
    version: 1,
    id: workflow.id,
    surface: 'outlook',
    mode: 'model',
    body:
      'Prepare a grounded response package for the active email. Read sufficient bounded source evidence, identify commitments and unresolved questions, and propose a reviewable draft. Check stated recipients, dates and requested attachments. Preserve the distinction between received item, open draft and a newly opened review form. Report the exact observed state.\n' +
      JSON.stringify({ instructions: workflow.instructions }),
    postconditions: [
      'Draft statements follow the captured evidence.',
      'Recipients and requested attachments are explicit.',
      'An opened review form is reported as opened.',
    ],
  };
}

/** Three-way refresh: unchanged generated output can advance; a manual edit is an explicit conflict. */
const canonical = (surface: Surface, value: string): string =>
  (surface === 'powerpoint' ? value.replace(/(?:\\r)?\\n/g, '\n') : value).replace(/\r\n?/g, '\n');

export async function prepareManagedTextRefresh(
  capture: TextTargetCapture,
  nextText: string,
  previousGeneratedText?: string,
): Promise<PreparedOfficeWorkflow> {
  const target = NativeTargetSchema.parse(capture.target);
  text.parse(nextText);
  if ((await contentHash({ text: capture.text })) !== target.revision)
    throw new Error('Refresh evidence does not match its target revision.');
  if (
    previousGeneratedText !== undefined &&
    canonical(target.surface, capture.text) !== canonical(target.surface, previousGeneratedText) &&
    canonical(target.surface, capture.text) !== canonical(target.surface, nextText)
  )
    throw new Error(
      'This output contains a manual edit. Resolve the conflict before refreshing it.',
    );
  const common = ` expectedDocumentId=${quote(target.documentId)} expectedTextHash=${quote(target.revision)} text=${quote(nextText)}`;
  let command: string;
  if (
    target.surface === 'powerpoint' &&
    target.scope === 'powerpoint-shape-text' &&
    target.parentObjectId
  )
    command =
      `/set-shape-text slideId=${quote(target.parentObjectId)} shapeId=${quote(target.objectId)}` +
      common;
  else if (target.surface === 'word' && target.scope === 'word-content-control')
    command = `/fill-content-control contentControlId=${quote(target.objectId)}` + common;
  else throw new Error('This target has no supported managed-text refresh capability.');
  return {
    version: 1,
    id: 'board-pack-refresh',
    surface: target.surface,
    mode: 'program',
    body:
      canonical(target.surface, capture.text) === canonical(target.surface, nextText)
        ? 'done "This output is already current."'
        : command + '\nfinish when=verified',
    postconditions: [
      'The exact captured object is updated.',
      'Manual edits are preserved or reported as conflicts.',
      'Native text readback matches the proposed output.',
    ],
  };
}

export const BoardPackBindingSchema = z
  .object({
    template: text,
    values: z
      .array(
        z
          .object({
            name: z.string().regex(/^[A-Za-z][A-Za-z0-9_]{0,63}$/),
            row: z.number().int().nonnegative(),
            column: z.number().int().nonnegative(),
          })
          .strict(),
      )
      .min(1)
      .max(100),
  })
  .strict();
/** Exact captured metrics; no rounding, invented numbers, or model inference during refresh. */
export function bindBoardPackMetrics(
  rawArtifact: unknown,
  rawBinding: z.input<typeof BoardPackBindingSchema>,
): {
  text: string;
  sources: z.infer<typeof TableArtifactSchema>['sources'];
  artifactId: string;
  artifactHash: string;
} {
  const artifact = TableArtifactSchema.parse(rawArtifact);
  const binding = BoardPackBindingSchema.parse(rawBinding);
  if (artifact.truncated) throw new Error('A board pack cannot use a truncated metric artifact.');
  const values = new Map<string, string>();
  for (const value of binding.values) {
    const row = artifact.rows[value.row];
    if (!row || value.column >= artifact.columns.length || values.has(value.name))
      throw new Error('Metric binding is missing, duplicated or out of bounds.');
    const cell = row[value.column];
    if (cell === null || cell === undefined)
      throw new Error('An explicitly bound metric is missing.');
    values.set(value.name, String(cell));
  }
  const rendered = binding.template.replace(
    /\{\{([A-Za-z][A-Za-z0-9_]*)\}\}/g,
    (_match, key: string) => {
      if (!values.has(key)) throw new Error('Template refers to an unbound metric.');
      return values.get(key)!;
    },
  );
  if (/\{\{|\}\}/.test(rendered)) throw new Error('Invalid or unresolved metric placeholder.');
  text.parse(rendered);
  return {
    text: rendered,
    sources: structuredClone(artifact.sources),
    artifactId: artifact.id,
    artifactHash: artifact.hash,
  };
}
