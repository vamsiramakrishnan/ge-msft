import { WordBridge } from '@ge/bridge-word';
import { ExcelBridge } from '@ge/bridge-excel';
import { PowerPointBridge } from '@ge/bridge-powerpoint';
import { OutlookBridge } from '@ge/bridge-outlook';
import { ArtifactStore } from '@ge/compute';
import { createNodeCompute } from '@ge/compute/node';
import {
  AssistSession,
  bindBoardPackMetrics,
  prepareManagedTextRefresh,
  type CommandLoopEvent,
  type DocBridge,
  type OfficeWorkflow,
  type RunCommandsOptions,
} from '@ge/runtime';
import type { SseEvent, TenantWorkflowId } from '@ge/contracts';
import type { StreamAssistClient } from '@ge/gemini-client';
import { installFakeExcel, excelSeed } from '../test-harness/fake-excel.js';
import { installFakeWord, wordSeed } from '../test-harness/fake-word.js';
import { installFakePowerPoint, powerPointSeed } from '../test-harness/fake-powerpoint.js';
import { installFakeOutlook, outlookSeed } from '../test-harness/fake-outlook.js';

export const TENANT_FLAGSHIP_MATRIX = [
  { id: 'contract-review', surface: 'word', modelRequired: true },
  { id: 'workbook-close', surface: 'excel', modelRequired: false },
  { id: 'board-pack-refresh', surface: 'powerpoint', modelRequired: false },
  { id: 'outlook-response', surface: 'outlook', modelRequired: true },
] as const;
type Event = SseEvent | CommandLoopEvent;
export interface TenantWorkflowFixture {
  session: AssistSession;
  bridge: DocBridge;
  run(options: RunCommandsOptions): AsyncGenerator<Event>;
  outcome(events: readonly Event[]): boolean;
  mutateSource?(): void;
  unchanged(): boolean;
  destinationUntouched(): boolean;
  dispose(): void;
}
const workflow: OfficeWorkflow = {
  id: 'workbook-close',
  inputs: {
    leftRange: 'Invoices!A1:C3',
    rightRange: 'Payments!A1:C4',
    leftCurrency: 2,
    rightCurrency: 2,
    tolerance: '0.001',
    destination: 'Report!A1',
  },
};
function receipts(events: readonly Event[]) {
  return events.flatMap((event) => (event.type === 'write-result' ? [event.result] : []));
}
/** Synthetic disposable documents. No native host/build certification can be produced from these fixtures. */
export async function tenantWorkflowFixture(
  id: TenantWorkflowId,
  client: StreamAssistClient,
): Promise<TenantWorkflowFixture> {
  let bridge: DocBridge;
  let restore: () => void;
  let run: (options: RunCommandsOptions) => AsyncGenerator<Event>;
  let outcome: (events: readonly Event[]) => boolean;
  let unchanged: () => boolean;
  let mutateSource: (() => void) | undefined;
  let snapshot: () => unknown;
  const compute = createNodeCompute();
  if (id === 'contract-review') {
    const sim = installFakeWord(
      wordSeed({
        selectionText: 'Payment is due in 45 days.',
        paragraphs: [
          { text: 'Vendor agreement', styleBuiltIn: 'Heading1' },
          { text: 'Payment is due in 45 days.', styleBuiltIn: 'Normal' },
          { text: 'Confidential information remains protected.', styleBuiltIn: 'Normal' },
        ],
      }),
    );
    bridge = new WordBridge();
    restore = sim.restore;
    snapshot = sim.snapshot;
    unchanged = () => sim.snapshot().bodyText.includes('Payment is due in 45 days.');
    outcome = (events) =>
      receipts(events).some(
        (result) =>
          result.ok &&
          result.kind === 'tracked-change' &&
          result.verification?.status === 'verified',
      ) &&
      receipts(events).length === 1 &&
      sim.snapshot().bodyText ===
        'Vendor agreement\nPayment is due in 30 days.\nConfidential information remains protected.';
    run = (options) =>
      session.runOfficeWorkflow(
        {
          id: 'contract-review',
          scope: 'Only the payment clause.',
          playbook:
            'Payment must be due in exactly 30 days. Propose one tracked replacement: Payment is due in 30 days. Preserve the confidentiality paragraph and all other content.',
        },
        options,
      );
  } else if (id === 'workbook-close') {
    const sim = installFakeExcel(
      excelSeed({
        sheets: [
          {
            name: 'Invoices',
            origin: 'A1',
            values: [
              ['key', 'amount', 'currency'],
              ['A', '0.30', 'USD'],
              ['B', '9007199254740993.01', 'USD'],
            ],
          },
          {
            name: 'Payments',
            origin: 'A1',
            values: [
              ['key', 'amount', 'currency'],
              ['A', '0.10', 'USD'],
              ['A', '0.20', 'USD'],
              ['B', '9007199254740993.00', 'USD'],
            ],
          },
          { name: 'Report', origin: 'A1', values: [['']] },
        ],
        selection: 'Invoices!A1:C3',
      }),
    );
    bridge = new ExcelBridge();
    restore = sim.restore;
    snapshot = sim.snapshot;
    unchanged = () =>
      sim.snapshot().sheets.find((sheet) => sheet.name === 'Report')!.values[0]?.[0] === '';
    mutateSource = () => {
      sim.seed.sheets.find((sheet) => sheet.name === 'Invoices')!.values[1]![1] = '9.99';
    };
    outcome = (events) =>
      receipts(events).some(
        (result) =>
          result.ok && result.kind === 'write-cells' && result.verification?.status === 'verified',
      ) &&
      JSON.stringify(
        sim.snapshot().sheets.find((sheet) => sheet.name === 'Report')!.values,
      ).includes('9007199254740993.01') &&
      JSON.stringify(
        sim.snapshot().sheets.find((sheet) => sheet.name === 'Report')!.values,
      ).includes('variance') &&
      JSON.stringify(
        sim.snapshot().sheets.find((sheet) => sheet.name === 'Report')!.values,
      ).includes('matched');
    run = (options) => session.runOfficeWorkflow(workflow, options);
  } else if (id === 'board-pack-refresh') {
    const sim = installFakePowerPoint(
      powerPointSeed({
        slides: [
          {
            id: 'slide-1',
            shapes: [
              { id: 'metric', text: 'Revenue: 100.00' },
              { id: 'manual', text: 'Manual risk note; preserve this.' },
            ],
          },
        ],
      }),
    );
    Object.defineProperty(Office.context.document, 'url', {
      value: 'https://synthetic.invalid/tenant-workflow.pptx',
      configurable: true,
    });
    bridge = new PowerPointBridge();
    restore = sim.restore;
    snapshot = sim.snapshot;
    unchanged = () =>
      sim.snapshot().slides[0]!.shapes.find((shape) => shape.id === 'metric')!.text ===
      'Revenue: 100.00';
    outcome = (events) =>
      receipts(events).some(
        (result) =>
          result.ok &&
          result.kind === 'set-shape-text' &&
          result.verification?.status === 'verified',
      ) &&
      sim.snapshot().slides[0]!.shapes.find((shape) => shape.id === 'metric')!.text ===
        'Revenue: 123456789012345678.90' &&
      sim.snapshot().slides[0]!.shapes.find((shape) => shape.id === 'manual')!.text ===
        'Manual risk note; preserve this.';
    run = async function* (options) {
      const artifact = await new ArtifactStore().add({
        title: 'Synthetic close metrics',
        labels: ['Revenue'],
        rows: [['123456789012345678.90']],
        sources: [],
        lineage: { parents: [], operation: 'query' },
      });
      const bound = bindBoardPackMetrics(artifact, {
        template: 'Revenue: {{revenue}}',
        values: [{ name: 'revenue', row: 0, column: 0 }],
      });
      const capture = await bridge.captureTextTarget!({
        id: 'pp:shape:slide-1:metric',
        title: 'Synthetic metric target',
        surface: 'powerpoint',
        kind: 'shape',
        hostRef: { type: 'powerpoint.shape', slideId: 'slide-1', shapeId: 'metric' },
      });
      const prepared = await prepareManagedTextRefresh(capture, bound.text, 'Revenue: 100.00');
      yield* session.runCommandProgram(prepared.body, options);
    };
  } else {
    const sim = installFakeOutlook(
      outlookSeed({
        itemId: 'synthetic-mail-1',
        subject: 'Launch checks',
        body: '<p>The security review is open. Owner Mira will provide an update by Friday. Do not invent a launch date.</p>',
        from: { emailAddress: 'sender@example.test', displayName: 'Synthetic Sender' },
      }),
    );
    bridge = new OutlookBridge();
    restore = sim.restore;
    snapshot = sim.snapshot;
    unchanged = () =>
      sim.snapshot().replyForms.length === 0 && sim.snapshot().newMessageForms.length === 0;
    outcome = (events) =>
      receipts(events).some(
        (result) =>
          result.ok &&
          ['reply-mail', 'create-mail'].includes(result.kind) &&
          result.completion === 'review-opened',
      ) &&
      receipts(events).length === 1 &&
      sim.snapshot().replyForms.length + sim.snapshot().newMessageForms.length === 1 &&
      [...sim.snapshot().replyForms, ...sim.snapshot().newMessageForms][0]!.htmlBody
        .replace(/<[^>]*>/g, ' ')
        .replace(/\s+/g, ' ')
        .trim() === 'The security review is open. Mira will update by Friday.';
    run = (options) =>
      session.runOfficeWorkflow(
        {
          id: 'outlook-response',
          instructions:
            'Prepare exactly one reviewable reply with this exact body: The security review is open. Mira will update by Friday. Do not invent any launch date, owner or commitment or add other body text. Open the form for review; do not send.',
        },
        options,
      );
  }
  const before = JSON.stringify(snapshot());
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: bridge.surface } },
    autoAttach: [],
    compute: async () => compute,
    context: { docState: true, lazyRead: true },
  });
  return {
    session,
    bridge,
    run,
    outcome: (events) =>
      outcome(events) &&
      session.executions.list().at(-1)?.status === 'completed' &&
      !events.some((event) => event.type === 'error' || event.type === 'capped'),
    destinationUntouched: unchanged,
    unchanged: () => unchanged() && JSON.stringify(snapshot()) === before,
    ...(mutateSource ? { mutateSource } : {}),
    dispose: () => {
      session.dispose();
      compute.dispose();
      restore();
    },
  };
}

/** These checks execute the shared runtime/approval/freshness boundary, never a synthetic success flag. */
export async function tenantWorkflowSafetyCanaries(
  client: StreamAssistClient,
): Promise<{ approvalRefusalVerified: boolean; sourceFreshnessVerified: boolean }> {
  const reject = await tenantWorkflowFixture('workbook-close', client);
  let approvalRefusalVerified = false;
  try {
    const events: Event[] = [];
    for await (const event of reject.run({ approvePlan: () => false })) events.push(event);
    approvalRefusalVerified =
      reject.unchanged() &&
      receipts(events).some(
        (result) =>
          !result.ok && ['unapproved', 'plan_unapproved'].includes(result.error?.code ?? ''),
      ) &&
      !receipts(events).some((result) => result.ok);
  } finally {
    reject.dispose();
  }
  const stale = await tenantWorkflowFixture('workbook-close', client);
  let sourceFreshnessVerified = false;
  try {
    const events: Event[] = [];
    for await (const event of stale.run({
      approvePlan: () => {
        stale.mutateSource!();
        return true;
      },
    }))
      events.push(event);
    sourceFreshnessVerified =
      stale.destinationUntouched() &&
      receipts(events).some(
        (result) =>
          !result.ok && ['source_conflict', 'stale_source'].includes(result.error?.code ?? ''),
      ) &&
      !receipts(events).some((result) => result.ok);
  } finally {
    stale.dispose();
  }
  return { approvalRefusalVerified, sourceFreshnessVerified };
}
