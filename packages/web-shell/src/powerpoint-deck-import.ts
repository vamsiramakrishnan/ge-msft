import {
  ActuationRequestSchema,
  DeckTemplateCaptureSchema,
  asChangeId,
  type ActuationRequest,
  type ChangeId,
  type ProvenancePayload,
} from '@ge/contracts';
import type {
  CompiledDeckArtifact,
  CompileDeckOptions,
  BoardPackTemplate,
  BoardPackLineage,
} from '@ge/deck-compiler';
import type { DeckSpec } from '@ge/contracts';

export interface BuildPowerPointDeckImportOptions {
  readonly deckSpec?: unknown;
  readonly boardPack?: {
    template: BoardPackTemplate;
    artifacts: readonly unknown[];
    currentSpec?: unknown;
    previousGeneratedSpec?: unknown;
  };
  readonly changeId?: ChangeId | string;
  readonly formatting?: 'KeepSourceFormatting' | 'UseDestinationTheme';
  readonly targetSlideId?: string;
  readonly provenance?: ProvenancePayload;
  readonly compile?: CompileDeckOptions;
}

export interface PowerPointDeckImportPlan {
  readonly artifact: CompiledDeckArtifact;
  readonly request: ActuationRequest;
  readonly boundSpec?: DeckSpec;
  readonly lineage?: readonly BoardPackLineage[];
}

export async function buildPowerPointDeckImportRequest(
  options: BuildPowerPointDeckImportOptions,
): Promise<PowerPointDeckImportPlan> {
  const { compileDeckSpecToBase64, compileBoardPack } = await import('@ge/deck-compiler');
  if ((options.deckSpec === undefined) === (options.boardPack === undefined))
    throw new Error('Provide exactly one deck specification or bound board pack.');
  const bound = options.boardPack
    ? await compileBoardPack(options.boardPack.template, options.boardPack.artifacts, {
        ...options.compile,
        currentSpec: options.boardPack.currentSpec,
        previousGeneratedSpec: options.boardPack.previousGeneratedSpec,
      })
    : undefined;
  const artifact =
    bound?.artifact ?? (await compileDeckSpecToBase64(options.deckSpec, options.compile));
  const candidate = {
    changeId:
      typeof options.changeId === 'string'
        ? asChangeId(options.changeId)
        : (options.changeId ?? asChangeId(`deck-${crypto.randomUUID()}`)),
    kind: 'insert-slide',
    surface: 'powerpoint',
    params: {
      deck: {
        format: 'pptx',
        base64: artifact.base64,
        slideCount: artifact.slideCount,
        specFingerprint: artifact.specFingerprint,
        ...(options.formatting ? { formatting: options.formatting } : {}),
        ...(options.targetSlideId ? { targetSlideId: options.targetSlideId } : {}),
      },
    },
    ...(options.provenance ? { provenance: options.provenance } : {}),
  };
  return {
    artifact,
    ...(bound ? { boundSpec: bound.spec, lineage: bound.lineage } : {}),
    request: ActuationRequestSchema.parse(candidate),
  };
}

/** Builds a normal reviewed effect from current native template evidence, never executes it. */
export function buildPowerPointLayoutRequest(
  rawCapture: unknown,
  options: {
    masterId: string;
    layoutId: string;
    slideId: string;
    changeId?: string;
    provenance?: ProvenancePayload;
  },
): ActuationRequest {
  const capture = DeckTemplateCaptureSchema.parse(rawCapture);
  if (!capture.complete)
    throw new Error('Capture a complete native template before applying a layout.');
  const layouts = capture.masters
    .filter((master) => master.id === options.masterId)
    .flatMap((master) => master.layouts.filter((layout) => layout.id === options.layoutId));
  if (layouts.length !== 1) throw new Error('Choose one captured native layout.');
  const layout = layouts[0]!;
  return ActuationRequestSchema.parse({
    surface: 'powerpoint',
    kind: 'apply-slide-layout',
    changeId: asChangeId(options.changeId ?? `layout-${crypto.randomUUID()}`),
    params: {
      target: { slideId: options.slideId, expectedDocumentId: capture.documentId },
      layout: {
        masterId: options.masterId,
        layoutId: layout.id,
        expectedLayoutHash: layout.revision,
      },
    },
    ...(options.provenance ? { provenance: options.provenance } : {}),
  });
}
