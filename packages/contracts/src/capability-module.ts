import { z } from 'zod';
import { ActuationParamsSchema, type ActuationKind } from './capability.js';
import type { Surface } from './context.js';

/** Parameter contracts describe admitted effects. Installed handlers retain execution authority. */
export interface CapabilityModule {
  surface: Surface;
  kind: ActuationKind;
  version: 1;
  params: z.ZodType;
  /** Required evidence; the bridge must report unknown when it cannot observe this condition. */
  postcondition: string;
  compensation: 'captured-inverse' | 'host-specific';
}
type ParamKey = keyof typeof ActuationParamsSchema.shape;

/** Reject ignored fields, including nested fields, rather than quietly laundering their intent. */
function strictFields(schema: z.ZodTypeAny): z.ZodTypeAny {
  if (schema instanceof z.ZodOptional) return strictFields(schema.unwrap()).optional();
  if (schema instanceof z.ZodNullable) return strictFields(schema.unwrap()).nullable();
  if (schema instanceof z.ZodDefault)
    return strictFields(schema.removeDefault()).default(schema._def.defaultValue);
  if (schema instanceof z.ZodObject)
    return z
      .object(
        Object.fromEntries(
          Object.entries(schema.shape).map(([key, value]) => [
            key,
            strictFields(value as z.ZodTypeAny),
          ]),
        ),
      )
      .strict();
  if (schema instanceof z.ZodArray)
    return new z.ZodArray({ ...schema._def, type: strictFields(schema.element) });
  if (schema instanceof z.ZodEffects)
    return new z.ZodEffects({ ...schema._def, schema: strictFields(schema.innerType()) });
  if (schema instanceof z.ZodDiscriminatedUnion) {
    const options = schema.options.map((option: z.ZodTypeAny) => strictFields(option));
    return z.discriminatedUnion(
      schema.discriminator,
      options as [
        z.ZodDiscriminatedUnionOption<string>,
        ...z.ZodDiscriminatedUnionOption<string>[],
      ],
    );
  }
  if (schema instanceof z.ZodUnion)
    return z.union(
      schema.options.map((option: z.ZodTypeAny) => strictFields(option)) as [
        z.ZodTypeAny,
        z.ZodTypeAny,
        ...z.ZodTypeAny[],
      ],
    );
  if (schema instanceof z.ZodRecord)
    return z.record(schema.keySchema, strictFields(schema.valueSchema));
  return schema;
}
const target = ActuationParamsSchema.shape.target.unwrap();
const wordTarget = target
  .pick({
    matchText: true,
    contextHint: true,
    contentControlId: true,
    commentId: true,
    expectedDocumentId: true,
    expectedTextHash: true,
  })
  .strict()
  .optional();
const excelTarget = target
  .pick({ range: true, commentId: true, expectedDocumentId: true })
  .strict()
  .optional();
const powerpointTarget = target
  .pick({
    slideIndex: true,
    slideId: true,
    shapeId: true,
    shapeName: true,
    expectedDocumentId: true,
    expectedTextHash: true,
  })
  .strict()
  .optional();
const outlookTarget = target.pick({ expectedDocumentId: true }).strict().optional();
const targets = {
  word: wordTarget,
  excel: excelTarget,
  powerpoint: powerpointTarget,
  outlook: outlookTarget,
};
function module(
  surface: keyof typeof targets,
  kind: ActuationKind,
  fields: readonly ParamKey[],
  postcondition: string,
  compensation: CapabilityModule['compensation'] = 'host-specific',
): CapabilityModule {
  const shape = Object.fromEntries(
    fields.map((key) => [
      key,
      key === 'target' ? targets[surface] : strictFields(ActuationParamsSchema.shape[key]),
    ]),
  );
  return {
    surface,
    kind,
    version: 1,
    params:
      kind === 'tracked-change'
        ? z.object(shape).strict().required({ text: true })
        : z.object(shape).strict(),
    postcondition,
    compensation,
  };
}
const modules: readonly CapabilityModule[] = [
  module(
    'word',
    'insert-text',
    ['text', 'textFont', 'position', 'target', 'sources'],
    'Read back inserted text at the resolved live range; unsupported readback remains unverified.',
  ),
  module(
    'word',
    'replace-selection',
    ['text', 'target', 'sources'],
    'Read back replacement text at the captured selection; preserve unrelated text.',
  ),
  module(
    'word',
    'insert-ooxml',
    ['ooxml', 'target', 'sources'],
    'Observe the inserted native object at the resolved range; callback success alone is unverified.',
  ),
  module(
    'word',
    'tracked-change',
    ['text', 'target', 'sources'],
    'Read back the inserted replacement at the resolved live range.',
  ),
  module(
    'word',
    'fill-content-control',
    ['text', 'target', 'sources'],
    'Read back the exact current-revision text in the captured content control.',
  ),
  module(
    'word',
    'add-comment',
    ['text', 'target', 'sources'],
    'Observe the host-minted comment and its resolved content anchor.',
  ),
  module(
    'word',
    'comment-reply',
    ['text', 'target', 'resolveComment', 'sources'],
    'Observe the reply on the named comment and requested resolution state.',
  ),
  module(
    'word',
    'apply-style',
    ['style', 'target', 'sources'],
    'Read back the requested style on the resolved live range.',
  ),
  module(
    'word',
    'insert-table',
    ['tableGrid', 'target', 'position', 'sources'],
    'Read back native table dimensions and cell values at the resolved insertion range.',
  ),
  module(
    'word',
    'insert-content-control',
    ['contentControl', 'target', 'sources'],
    'Observe the host-minted control identity, type and requested tag/title.',
  ),
  module(
    'word',
    'insert-hyperlink',
    ['hyperlink', 'target', 'sources'],
    'Observe the screened hyperlink address on the resolved live range.',
  ),
  module(
    'word',
    'find-replace',
    ['findReplace', 'sources'],
    'Observe replacements only at the approved matched ranges; preserve other text.',
  ),
  module(
    'excel',
    'write-cells',
    ['cells', 'cellValues', 'cellFormulas', 'target', 'sources'],
    'Fresh values/formulas match the requested rectangular range.',
    'captured-inverse',
  ),
  module(
    'excel',
    'format-cells',
    ['format', 'target', 'sources'],
    'Read back the requested formatting on the explicit range.',
  ),
  module(
    'excel',
    'sort-range',
    ['sortRange', 'target', 'sources'],
    'Fresh rows preserve the source multiset and satisfy the requested ordering.',
    'captured-inverse',
  ),
  module(
    'excel',
    'create-table',
    ['table', 'target', 'sources'],
    'Observe the host-minted native table identity, exact range and header configuration.',
  ),
  module(
    'excel',
    'insert-chart',
    ['chart', 'target', 'sources'],
    'Observe the host-minted native chart, chart type and source ranges.',
  ),
  module(
    'excel',
    'format-conditional',
    ['conditional', 'target', 'sources'],
    'Observe the host-minted rule on the exact approved range.',
  ),
  module(
    'excel',
    'add-comment',
    ['text', 'target', 'sources'],
    'Observe the host-minted comment at the exact approved cell.',
  ),
  module(
    'excel',
    'comment-reply',
    ['text', 'target', 'resolveComment', 'sources'],
    'Observe the reply on the named comment and requested resolution state.',
  ),
  module(
    'excel',
    'insert-pivot',
    ['pivot', 'target', 'sources'],
    'Read back the native pivot identity, source, destination and configured hierarchy fields.',
  ),
  module(
    'excel',
    'filter-range',
    ['filterRange', 'target', 'sources'],
    'Read back the native filter range, column and criteria without changing cell values.',
  ),
  module(
    'excel',
    'set-data-validation',
    ['validation', 'target', 'sources'],
    'Read back the native literal-list rule and Stop error alert on the exact captured rectangle.',
  ),
  module(
    'excel',
    'manage-worksheet',
    ['worksheet', 'target', 'sources'],
    'Read back the exact native worksheet identity and requested lifecycle state.',
  ),
  module(
    'powerpoint',
    'insert-slide',
    ['slide', 'deck', 'target', 'text', 'ooxml', 'sources'],
    'Read back newly minted slide IDs and requested native text/placement.',
  ),
  module(
    'powerpoint',
    'set-shape-text',
    ['text', 'target', 'sources'],
    'Read back the requested text at the captured slide and shape identity.',
  ),
  module(
    'powerpoint',
    'add-shape',
    ['shape', 'target', 'sources'],
    'Observe the host-minted shape, text and requested geometry on the named slide.',
  ),
  module(
    'powerpoint',
    'format-shape',
    ['shapeFormat', 'target', 'sources'],
    'Read back requested format properties on the named slide/shape identity.',
  ),
  module(
    'powerpoint',
    'add-table-slide',
    ['tableGrid', 'slide', 'target', 'sources'],
    'Observe the host-minted native table, cell values and bounded placement.',
  ),
  module(
    'powerpoint',
    'apply-slide-layout',
    ['layout', 'target', 'sources'],
    'Read back the requested existing layout on the exact captured slide.',
  ),
  module(
    'outlook',
    'reply-mail',
    ['mail', 'text', 'attachment', 'target', 'sources'],
    'Observe a reviewable reply form opening; never imply sending or persisted draft delivery.',
  ),
  module(
    'outlook',
    'create-mail',
    ['mail', 'text', 'target', 'sources'],
    'Observe a reviewable new-message form opening; never imply sending.',
  ),
  module(
    'outlook',
    'set-recipients',
    ['mail', 'target', 'sources'],
    'Read back the requested recipient fields on the same active compose item.',
  ),
  module(
    'outlook',
    'add-attachment',
    ['attachment', 'target', 'sources'],
    'Observe the host-minted attachment on the same active compose item.',
  ),
  module(
    'outlook',
    'set-body',
    ['mail', 'text', 'html', 'target', 'sources'],
    'Read back the requested body on the same active compose item.',
  ),
  module(
    'outlook',
    'set-subject',
    ['mail', 'target', 'sources'],
    'Read back the requested subject on the same active compose item.',
  ),
];
export function capabilityModule(
  surface: Surface,
  kind: ActuationKind,
): CapabilityModule | undefined {
  return modules.find((entry) => entry.surface === surface && entry.kind === kind);
}
export function listCapabilityModules(): readonly CapabilityModule[] {
  return modules.map((entry) => ({ ...entry }));
}
