import { z } from 'zod';
import { ActuationParamsSchema, type ActuationKind } from './capability.js';
import type { Surface } from './context.js';

/** Incremental migration: installed bridge handlers retain all execution authority. */
export interface CapabilityModule {
  surface: Surface;
  kind: ActuationKind;
  version: 1;
  params: z.ZodType;
  postcondition: string;
  compensation: 'captured-inverse' | 'host-specific';
}
const target = ActuationParamsSchema.shape.target.unwrap();
const modules: readonly CapabilityModule[] = [
  {
    surface: 'word',
    kind: 'tracked-change',
    version: 1,
    params: z
      .object({
        text: z.string(),
        target: target
          .pick({ matchText: true, contextHint: true, contentControlId: true })
          .strict()
          .optional(),
        sources: ActuationParamsSchema.shape.sources,
      })
      .strict(),
    postcondition: 'Read back the inserted replacement at the resolved live range.',
    compensation: 'host-specific',
  },
  {
    surface: 'excel',
    kind: 'write-cells',
    version: 1,
    params: ActuationParamsSchema.pick({
      cells: true,
      cellValues: true,
      cellFormulas: true,
      target: true,
      sources: true,
    }).strict(),
    postcondition: 'Fresh values/formulas match the requested rectangular range.',
    compensation: 'captured-inverse',
  },
  {
    surface: 'excel',
    kind: 'sort-range',
    version: 1,
    params: ActuationParamsSchema.pick({ sortRange: true, target: true }).strict(),
    postcondition: 'Fresh rows preserve the source multiset and satisfy the requested ordering.',
    compensation: 'captured-inverse',
  },
  {
    surface: 'powerpoint',
    kind: 'insert-slide',
    version: 1,
    params: ActuationParamsSchema.pick({
      slide: true,
      deck: true,
      target: true,
      text: true,
      ooxml: true,
    }).strict(),
    postcondition: 'Read back the newly minted slide IDs and requested placement.',
    compensation: 'host-specific',
  },
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
