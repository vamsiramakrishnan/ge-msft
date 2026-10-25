import { z } from 'zod';
import { DeckSpecSchema, TableArtifactSchema, type DeckSpec } from '@ge/contracts';
import {
  compileDeckSpecToBase64,
  type CompiledDeckArtifact,
  type CompileDeckOptions,
} from './index.js';

const cell = z
  .object({
    artifactId: z.string().min(1).max(256),
    row: z.number().int().nonnegative(),
    column: z.number().int().nonnegative(),
  })
  .strict();
export const BoardPackTemplateSchema = z
  .object({
    version: z.literal(1),
    spec: DeckSpecSchema,
    bindings: z
      .array(
        z.discriminatedUnion('kind', [
          z
            .object({
              kind: z.literal('text'),
              slideId: z.string().min(1).max(120),
              element: z.number().int().nonnegative(),
              values: z
                .record(z.string().regex(/^[A-Za-z][A-Za-z0-9_]{0,63}$/), cell)
                .refine(
                  (values) => Object.keys(values).length <= 100,
                  'At most 100 metric names per text binding.',
                ),
            })
            .strict(),
          z
            .object({
              kind: z.literal('chart'),
              slideId: z.string().min(1).max(120),
              element: z.number().int().nonnegative(),
              artifactId: z.string().min(1).max(256),
              rows: z.array(z.number().int().nonnegative()).min(1).max(500),
              labelColumn: z.number().int().nonnegative(),
              series: z
                .array(
                  z
                    .object({
                      name: z.string().min(1).max(200),
                      column: z.number().int().nonnegative(),
                    })
                    .strict(),
                )
                .min(1)
                .max(12),
            })
            .strict(),
        ]),
      )
      .max(100),
  })
  .strict();
export type BoardPackTemplate = z.input<typeof BoardPackTemplateSchema>;
export interface BoardPackLineage {
  artifactId: string;
  artifactHash: string;
  sources: z.infer<typeof TableArtifactSchema>['sources'];
  cells: Array<{ row: number; column: number; exactValue: string }>;
}
export interface BoundBoardPack {
  spec: DeckSpec;
  lineage: BoardPackLineage[];
}

/** Bind only explicitly selected cells; exact decimal text is retained alongside chart numeric data. */
export function bindBoardPackTemplate(
  raw: BoardPackTemplate,
  inputs: readonly unknown[],
  currentSpec?: unknown,
  previousGeneratedSpec?: unknown,
): BoundBoardPack {
  if (inputs.length > 16) throw new Error('A board pack accepts at most 16 metric artifacts.');
  const template = BoardPackTemplateSchema.parse(raw);
  const spec = structuredClone(
    currentSpec === undefined ? template.spec : DeckSpecSchema.parse(currentSpec),
  );
  const previous =
    previousGeneratedSpec === undefined ? undefined : DeckSpecSchema.parse(previousGeneratedSpec);
  if ((currentSpec === undefined) !== (previousGeneratedSpec === undefined))
    throw new Error('Managed refresh needs both current and previous generated specs.');
  const artifacts = new Map<string, z.infer<typeof TableArtifactSchema>>();
  let artifactBytes = 0;
  for (const input of inputs) {
    const artifact = TableArtifactSchema.parse(input);
    artifactBytes += new TextEncoder().encode(JSON.stringify(artifact)).byteLength;
    if (artifactBytes > 8 * 1024 * 1024)
      throw new Error('Board-pack metric artifacts exceed the 8 MiB aggregate budget.');
    if (artifact.truncated || artifacts.has(artifact.id))
      throw new Error('Board-pack artifacts must be complete and uniquely identified.');
    artifacts.set(artifact.id, artifact);
  }
  const lineage = new Map<string, BoardPackLineage>();
  const used = new Set<string>();
  const recordedCells = new Set<string>();
  let bindingReads = 0;
  const value = (id: string, row: number, column: number): string => {
    if (++bindingReads > 20000)
      throw new Error('Board-pack bindings exceed the 20,000-cell read budget.');
    const artifact = artifacts.get(id);
    const observed = artifact?.rows[row]?.[column];
    if (
      !artifact ||
      column >= artifact.columns.length ||
      observed === undefined ||
      observed === null
    )
      throw new Error('A bound board-pack cell is missing or outside the artifact.');
    let record = lineage.get(id);
    if (!record) {
      record = {
        artifactId: id,
        artifactHash: artifact.hash,
        sources: structuredClone(artifact.sources),
        cells: [],
      };
      lineage.set(id, record);
    }
    const exactValue = String(observed);
    const cellKey = `${id}:${row}:${column}`;
    if (!recordedCells.has(cellKey)) {
      recordedCells.add(cellKey);
      record.cells.push({ row, column, exactValue });
    }
    return exactValue;
  };
  for (const binding of template.bindings) {
    const key = `${binding.slideId}:${binding.element}`;
    if (used.has(key)) throw new Error('A board-pack element may have only one binding.');
    used.add(key);
    const findSlide = (deck: DeckSpec) => {
      const found = deck.slides.filter((slide) => slide.id === binding.slideId);
      if (found.length !== 1) throw new Error('A binding needs exactly one stable slide id.');
      return found[0]!;
    };
    const original = findSlide(template.spec).elements[binding.element];
    const destination = findSlide(spec);
    const observed = destination.elements[binding.element];
    if (!original || !observed) throw new Error('A board-pack element is missing.');
    if (previous) {
      const prior = findSlide(previous).elements[binding.element];
      if (JSON.stringify(prior) !== JSON.stringify(observed))
        throw new Error(
          'A managed board-pack element contains manual edits. Resolve the conflict before refresh.',
        );
    }
    if (binding.kind === 'text') {
      if (original.kind !== 'text' || observed.kind !== 'text')
        throw new Error('Text binding needs a text element.');
      const names = new Set<string>();
      const text = original.text.replace(
        /\{\{([A-Za-z][A-Za-z0-9_]{0,63})\}\}/g,
        (_match, name: string) => {
          names.add(name);
          const source = binding.values[name];
          if (!source) throw new Error('Unbound board-pack text placeholder.');
          return value(source.artifactId, source.row, source.column);
        },
      );
      if (/\{\{|\}\}/.test(text) || Object.keys(binding.values).some((name) => !names.has(name)))
        throw new Error('Invalid or unused board-pack text binding.');
      destination.elements[binding.element] = { ...observed, text };
    } else {
      if (original.kind !== 'chart' || observed.kind !== 'chart')
        throw new Error('Chart binding needs an editable chart element.');
      const labels = binding.rows.map((row) => value(binding.artifactId, row, binding.labelColumn));
      const series = binding.series.map((series) => ({
        name: series.name,
        values: binding.rows.map((row) =>
          chartNumber(value(binding.artifactId, row, series.column)),
        ),
      }));
      destination.elements[binding.element] = { ...observed, labels, series };
    }
  }
  return { spec: DeckSpecSchema.parse(spec), lineage: [...lineage.values()] };
}

/** Compile/refresh the artifact; host insertion still uses the ordinary reviewed insert-slide request. */
export async function compileBoardPack(
  raw: BoardPackTemplate,
  inputs: readonly unknown[],
  options: CompileDeckOptions & { currentSpec?: unknown; previousGeneratedSpec?: unknown } = {},
): Promise<BoundBoardPack & { artifact: CompiledDeckArtifact }> {
  const bound = bindBoardPackTemplate(
    raw,
    inputs,
    options.currentSpec,
    options.previousGeneratedSpec,
  );
  const artifact = await compileDeckSpecToBase64(bound.spec, options);
  return { ...bound, artifact };
}

function decimal(value: string): string {
  const match = /^(-?)(\d+)(?:\.(\d*))?(?:e([+-]?\d+))?$/i.exec(value);
  if (!match) throw new Error('Chart metrics require decimal values.');
  const exponent = Number(match[4] ?? 0) - (match[3]?.length ?? 0);
  let digits = `${match[2]}${match[3] ?? ''}`.replace(/^0+/, '') || '0';
  let scale = exponent;
  while (digits.endsWith('0') && digits !== '0') {
    digits = digits.slice(0, -1);
    scale++;
  }
  return digits === '0' ? '0' : `${match[1]}${digits}e${scale}`;
}
function chartNumber(exact: string): number {
  if (!/^-?\d+(?:\.\d+)?$/.test(exact))
    throw new Error('Chart metrics require exact decimal text.');
  const number = Number(exact);
  if (
    !Number.isFinite(number) ||
    decimal(exact) !== decimal(String(number)) ||
    (Number.isInteger(number) && !Number.isSafeInteger(number))
  )
    throw new Error(
      'This metric cannot be represented safely by a native chart; use a labelled exact-text value.',
    );
  return number;
}
