import {
  OfficeApplicationBindingSchema,
  OfficeHandoffPlanSchema,
  TableArtifactSchema,
  contentHash,
  assessActuationResult,
  type ContextRef,
  type DurableCommandProgram,
  type OfficeApplicationBinding,
  type OfficeHandoffPlan,
  type TableArtifact,
  type TextTargetCapture,
} from '@ge/contracts';
import type { AssistSession, CommandLoopEvent, RunCommandsOptions } from './assist-session.js';
import type { SseEvent } from '@ge/contracts';
import type { DurableProgramEvent, DurableStageCheckpoint } from './durable-program.js';
import { bindBoardPackMetrics, prepareManagedTextRefresh } from './office-workflows.js';

type SessionEvent = SseEvent | CommandLoopEvent | DurableProgramEvent;
export interface OfficeApplicationPort {
  binding(): Promise<OfficeApplicationBinding>;
  artifact(id: string): Promise<TableArtifact>;
  capture(target: TextTargetCapture['target']): Promise<TextTargetCapture>;
  run(
    program: DurableCommandProgram,
    options: RunCommandsOptions & {
      revalidate?: (
        program: DurableCommandProgram,
        checkpoints: DurableStageCheckpoint[],
      ) => Promise<boolean>;
    },
  ): AsyncGenerator<SessionEvent>;
}
/** The composition root supplies a current authenticated object ID; no token is passed or retained. */
export function createOfficeApplicationPort(
  session: Pick<
    AssistSession,
    | 'applicationContext'
    | 'exportAnalysisArtifact'
    | 'captureNativeTextTarget'
    | 'runDurableProgram'
  >,
  expectedObjectId: string | undefined,
  currentObjectId: () => Promise<string | undefined>,
): OfficeApplicationPort {
  return {
    async binding() {
      const id = await currentObjectId();
      if (!id || !expectedObjectId)
        throw new Error('Cross-application work needs a current authenticated object ID.');
      if (id !== expectedObjectId)
        throw new Error(
          'The authenticated account changed. Rebuild the Office session before creating a handoff.',
        );
      const context = await session.applicationContext();
      return OfficeApplicationBindingSchema.parse({
        ...context,
        principalHash: await contentHash({ provider: 'entra', objectId: id }),
      });
    },
    artifact: (id) => session.exportAnalysisArtifact(id),
    capture: (target) => {
      const ref: ContextRef =
        target.surface === 'word'
          ? {
              id: target.locator,
              surface: 'word',
              kind: 'paragraph',
              title: 'Managed output',
              hostRef: { type: 'word.contentControl', contentControlId: target.objectId },
            }
          : {
              id: target.locator,
              surface: 'powerpoint',
              kind: 'shape',
              title: 'Managed output',
              hostRef: {
                type: 'powerpoint.shape',
                slideId: target.parentObjectId ?? '',
                shapeId: target.objectId,
              },
            };
      return session.captureNativeTextTarget(ref);
    },
    run: (program, options) => session.runDurableProgram(program, options),
  };
}
export type OfficeHandoffEvent =
  | { type: 'handoff-event'; stepId: string; event: SessionEvent }
  | {
      type: 'handoff-step';
      stepId: string;
      status: 'verified' | 'current' | 'review-required' | 'blocked';
      artifactHash: string;
    };

const same = (left: OfficeApplicationBinding, right: OfficeApplicationBinding): boolean =>
  left.surface === right.surface &&
  left.documentId === right.documentId &&
  left.principalHash === right.principalHash;
const quote = (value: string): string =>
  JSON.stringify(value)
    .replace(/`/g, '\\u0060')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
const canonical = (text: string): string => text.replace(/\r\n?/g, '\n');

/** Typed coordination over real host sessions. Every destination retains its own approval and journal. */
export class OfficeWorkflowCoordinator {
  private readonly applications = new Map<string, OfficeApplicationPort>();
  private running = false;
  register(id: string, port: OfficeApplicationPort): () => void {
    if (
      !/^[a-z][a-z0-9-]{0,63}$/.test(id) ||
      this.applications.has(id) ||
      this.applications.size >= 8
    )
      throw new Error('Use a unique application ID within the eight-session limit.');
    this.applications.set(id, port);
    return () => {
      if (this.applications.get(id) === port) this.applications.delete(id);
    };
  }
  private application(id: string): OfficeApplicationPort {
    const port = this.applications.get(id);
    if (!port) throw new Error(`Open and register the required Office application: ${id}.`);
    return port;
  }
  private async fresh(plan: OfficeHandoffPlan): Promise<TableArtifact> {
    const source = this.application(plan.sourceApplication);
    if (!same(OfficeApplicationBindingSchema.parse(await source.binding()), plan.sourceBinding))
      throw new Error('The source identity or document changed. Review a new handoff.');
    // artifact() rechecks all native source versions in the originating session.
    const artifact = TableArtifactSchema.parse(await source.artifact(plan.artifactId));
    if (!same(OfficeApplicationBindingSchema.parse(await source.binding()), plan.sourceBinding))
      throw new Error(
        'The source identity or document changed during capture. Review a new handoff.',
      );
    if (new TextEncoder().encode(JSON.stringify(artifact)).byteLength > 1024 * 1024)
      throw new Error('A cross-application artifact must fit the one MiB handoff budget.');
    const hash = await contentHash({
      columns: artifact.columns,
      rows: artifact.rows,
      sources: artifact.sources,
      lineage: artifact.lineage,
      truncated: artifact.truncated,
    });
    if (
      artifact.truncated ||
      !artifact.sources.length ||
      artifact.hash !== hash ||
      artifact.id !== `a_${hash.slice(7, 31)}` ||
      hash !== plan.artifactHash
    )
      throw new Error(
        'The handoff artifact is incomplete or no longer matches its reviewed content.',
      );
    return artifact;
  }
  async *run(
    input: OfficeHandoffPlan,
    options: RunCommandsOptions = {},
  ): AsyncGenerator<OfficeHandoffEvent> {
    if (this.running) throw new Error('This coordinator is already running a handoff.');
    const plan = OfficeHandoffPlanSchema.parse(input);
    this.running = true;
    try {
      if (plan.sourceBinding.surface !== 'excel')
        throw new Error('Metric handoffs need an Excel source session.');
      // Validate every destination before the first effect; no partial workflow caused by a missing session.
      for (const step of plan.steps) {
        const current = OfficeApplicationBindingSchema.parse(
          await this.application(step.application).binding(),
        );
        if (
          !same(current, step.binding) ||
          current.principalHash !== plan.sourceBinding.principalHash
        )
          throw new Error(
            'Handoffs require the same current principal and the reviewed destination documents.',
          );
        if (
          step.kind === 'managed-text' &&
          (step.capture.target.documentId !== current.documentId ||
            step.capture.target.surface !== current.surface)
        )
          throw new Error('The managed target belongs to a different application or document.');
        if (step.kind === 'outlook-review' && current.surface !== 'outlook')
          throw new Error('Open the review package in Outlook.');
      }
      for (const step of plan.steps) {
        options.signal?.throwIfAborted();
        const destination = this.application(step.application);
        const artifact = await this.fresh(plan);
        const text = bindBoardPackMetrics(artifact, step.metrics).text;
        const guard = async (): Promise<void> => {
          options.signal?.throwIfAborted();
          await this.fresh(plan);
          if (
            !same(OfficeApplicationBindingSchema.parse(await destination.binding()), step.binding)
          )
            throw new Error('The destination identity or document changed. Review a new handoff.');
          await options.beforeDispatch?.();
        };
        await guard();
        const prepared =
          step.kind === 'managed-text'
            ? await prepareManagedTextRefresh(step.capture, text, step.previousGeneratedText)
            : {
                body: `compose ${quote(step.subject)} ${quote(text)}\ndone "Review the opened Outlook form before sending."`,
              };
        if (step.kind === 'managed-text' && canonical(step.capture.text) === canonical(text)) {
          const live = await destination.capture(step.capture.target);
          if (
            live.target.documentId !== step.binding.documentId ||
            live.target.objectId !== step.capture.target.objectId ||
            canonical(live.text) !== canonical(text)
          )
            throw new Error('The previously current managed output changed. Review it again.');
          yield {
            type: 'handoff-step',
            stepId: step.id,
            status: 'current',
            artifactHash: artifact.hash,
          };
          continue;
        }
        let verified = false;
        let reviewOpened = false;
        for await (const event of destination.run(
          {
            version: 1,
            id: `${plan.id}-${step.id}`,
            surface: step.binding.surface,
            stages: [{ id: 'apply', dependsOn: [], program: prepared.body }],
          },
          {
            ...options,
            beforeDispatch: guard,
            revalidate: async () => {
              await guard();
              if (step.kind !== 'managed-text') return false;
              const live = await destination.capture(step.capture.target);
              return (
                live.target.documentId === step.binding.documentId &&
                live.target.objectId === step.capture.target.objectId &&
                live.target.parentObjectId === step.capture.target.parentObjectId &&
                canonical(live.text) === canonical(text)
              );
            },
          },
        )) {
          if (event.type === 'write-result') {
            verified =
              assessActuationResult(event.result) === 'verified' && !event.result.recoveryPending;
            reviewOpened =
              event.result.ok &&
              event.result.completion === 'review-opened' &&
              !event.result.recoveryPending;
          }
          if (event.type === 'durable-progress' && event.state === 'resuming') verified = true;
          yield { type: 'handoff-event', stepId: step.id, event };
        }
        yield {
          type: 'handoff-step',
          stepId: step.id,
          status: verified ? 'verified' : reviewOpened ? 'review-required' : 'blocked',
          artifactHash: artifact.hash,
        };
        if (!verified) return;
      }
    } finally {
      this.running = false;
    }
  }
}
