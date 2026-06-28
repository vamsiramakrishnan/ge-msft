import {
  ActuationRequestSchema,
  asChangeId,
  type ActuationRequest,
  type ChangeId,
  type ProvenancePayload,
} from '@ge/contracts';
import type { CompiledDeckArtifact, CompileDeckOptions } from '@ge/deck-compiler';

export interface BuildPowerPointDeckImportOptions {
  readonly deckSpec: unknown;
  readonly changeId?: ChangeId | string;
  readonly formatting?: 'KeepSourceFormatting' | 'UseDestinationTheme';
  readonly targetSlideId?: string;
  readonly provenance?: ProvenancePayload;
  readonly compile?: CompileDeckOptions;
}

export interface PowerPointDeckImportPlan {
  readonly artifact: CompiledDeckArtifact;
  readonly request: ActuationRequest;
}

export async function buildPowerPointDeckImportRequest(
  options: BuildPowerPointDeckImportOptions,
): Promise<PowerPointDeckImportPlan> {
  const { compileDeckSpecToBase64 } = await import('@ge/deck-compiler');
  const artifact = await compileDeckSpecToBase64(options.deckSpec, options.compile);
  const candidate = {
    changeId:
      typeof options.changeId === 'string'
        ? asChangeId(options.changeId)
        : (options.changeId ?? asChangeId(`deck-${artifact.specFingerprint}`)),
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
    request: ActuationRequestSchema.parse(candidate),
  };
}
