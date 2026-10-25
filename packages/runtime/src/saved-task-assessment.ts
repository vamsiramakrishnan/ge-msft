import type { DocBridge } from './bridge.js';
import type { DurableProgramRecord, DurableStageCheckpoint } from './durable-program.js';
import { inspectDurableOutput, validateDurableSources } from './durable-inspection.js';
import { contentHash, type ActuationRequest, type ContextRef } from '@ge/contracts';
import { boundedNativeRead, type NativeReadOptions } from './bounded-native-read.js';

/** Pin the exact managed object observed for the UI preview; supplied pins are never rebased. */
export async function prepareInspectedTextTarget(
  input: ActuationRequest,
  bridge: DocBridge,
  options: NativeReadOptions = {},
): Promise<ActuationRequest> {
  if (!['fill-content-control', 'set-shape-text'].includes(input.kind)) return input;
  const target = input.params.target;
  const word =
    input.surface === 'word' && input.kind === 'fill-content-control' && target?.contentControlId;
  const powerpoint =
    input.surface === 'powerpoint' &&
    input.kind === 'set-shape-text' &&
    target?.slideId &&
    target?.shapeId;
  const key = bridge.documentKey?.();
  if (!target || (!word && !powerpoint) || !key || !bridge.captureTextTarget)
    throw new Error('This saved text task needs a stable native object and current read port.');
  const ref: ContextRef = word
    ? {
        id: `word:cc:${target.contentControlId}`,
        surface: 'word',
        kind: 'paragraph',
        title: 'Saved task target',
        hostRef: { type: 'word.contentControl', contentControlId: target.contentControlId! },
      }
    : {
        id: `pp:shape:${target.slideId}:${target.shapeId}`,
        surface: 'powerpoint',
        kind: 'shape',
        title: 'Saved task target',
        hostRef: { type: 'powerpoint.shape', slideId: target.slideId!, shapeId: target.shapeId! },
      };
  const captured = await boundedNativeRead(() => bridge.captureTextTarget!(ref), options);
  const documentId = await contentHash({ surface: bridge.surface, documentKey: key });
  const expectedId = word ? target.contentControlId : target.shapeId;
  const expectedParent = word ? undefined : target.slideId;
  if (
    bridge.surface !== input.surface ||
    bridge.documentKey?.() !== key ||
    captured.target.surface !== input.surface ||
    captured.target.documentId !== documentId ||
    captured.target.objectId !== expectedId ||
    captured.target.parentObjectId !== expectedParent ||
    captured.target.scope !== (word ? 'word-content-control' : 'powerpoint-shape-text') ||
    new TextEncoder().encode(captured.text).byteLength > 64 * 1024 ||
    captured.target.revision !== (await contentHash({ text: captured.text }))
  )
    throw new Error('Native text capture did not match the saved target.');
  if (
    (target.expectedDocumentId && target.expectedDocumentId !== documentId) ||
    (target.expectedTextHash && target.expectedTextHash !== captured.target.revision)
  )
    throw new Error(
      'The supplied saved text target changed. Review a new task rather than rebasing its pin.',
    );
  return {
    ...input,
    params: {
      ...input.params,
      target: {
        ...target,
        expectedDocumentId: documentId,
        expectedTextHash: captured.target.revision,
      },
    },
  };
}

export interface DurableStageAssessment {
  id: string;
  status: 'current' | 'pending' | 'blocked' | 'unsupported' | 'changed';
  reason?: string;
}
export interface DurableProgramAssessment {
  record: DurableProgramRecord;
  canResume: boolean;
  blockers: string[];
  stages: DurableStageAssessment[];
}

/** This is read-only evidence, not an approval, a reconciliation decision or retry authority. */
export async function assessDurableCheckpoints(
  checkpoints: readonly DurableStageCheckpoint[],
  bridge: DocBridge,
  readOptions: NativeReadOptions = {},
): Promise<DurableStageAssessment[]> {
  const options = { deadline: Date.now() + 10000, ...readOptions };
  const stages: DurableStageAssessment[] = [];
  for (const checkpoint of checkpoints) {
    if (checkpoint.state === 'pending') {
      stages.push({ id: checkpoint.id, status: 'pending' });
    } else if (checkpoint.state !== 'verified') {
      stages.push({
        id: checkpoint.id,
        status: 'blocked',
        reason: 'An interrupted or unverified effect needs explicit native inspection.',
      });
    } else {
      const assessment = await inspectDurableOutput(
        checkpoint.inspection,
        checkpoint.receipt,
        bridge,
        options,
      );
      stages.push({ id: checkpoint.id, ...assessment });
    }
  }
  return stages;
}

export async function assessSavedTask(
  record: DurableProgramRecord,
  bridge: DocBridge,
  readOptions: NativeReadOptions = {},
): Promise<DurableProgramAssessment> {
  const options = { deadline: Date.now() + 10000, ...readOptions };
  const blockers: string[] = [];
  if (record.archivedAt)
    blockers.push('This completed task is hidden; its task ID cannot be reused.');
  if (record.runId) blockers.push('Another pane owns this task or its execution was interrupted.');
  const sources = record.archivedAt
    ? undefined
    : await validateDurableSources(record.program, bridge, options);
  if (sources && sources.status !== 'current')
    blockers.push(sources.reason ?? 'Task inputs could not be revalidated.');
  const stages = await assessDurableCheckpoints(record.stages, bridge, options);
  for (const stage of stages)
    if (!['current', 'pending'].includes(stage.status))
      blockers.push(stage.reason ?? `Stage ${stage.id} is not current.`);
  return {
    record: structuredClone(record),
    canResume: blockers.length === 0,
    blockers: [...new Set(blockers)],
    stages,
  };
}
