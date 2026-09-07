import {
  asChangeId,
  approvalClassOf,
  isReversibleKind,
  commandPaletteFor,
  type ActuationRequest,
  type ContextRef,
  type Surface,
} from '@ge/contracts';
import { PanelController, type AssistLike } from '../controller.js';
import type { RunCommandsOptions } from '@ge/runtime';
import { renderCommandLine } from '../render-command.js';

const SAMPLE: Record<
  Surface,
  {
    title: string;
    body: string;
    answer: string;
    kind: ActuationRequest['kind'];
    params: ActuationRequest['params'];
  }
> = {
  word: {
    title: 'Launch decision.docx',
    body: 'The pilot is complete. Two integration risks remain.',
    answer:
      '## Decision brief\nApprove the next pilot phase, subject to the two outstanding integration checks.\n\n- **Evidence:** the pilot completed its agreed scope.\n- **Open question:** who owns the remaining checks?\n- **Next step:** assign owners before committing the launch date.',
    kind: 'tracked-change',
    params: {
      target: { matchText: 'The pilot is complete.' },
      text: 'The pilot met its agreed scope. Complete the integration checks before approving launch.',
    },
  },
  excel: {
    title: 'Sales!A1:D12',
    body: 'Region\tActual\tForecast\nAPAC\t84\t100\nEMEA\t108\t100',
    answer:
      '## Actuals against forecast\n\n| Region | Actual | Forecast | Variance |\n| --- | --- | --- | --- |\n| APAC | 84 | 100 | -16% |\n| EMEA | 108 | 100 | +8% |\n\nAPAC accounts for the shortfall. Check timing and reporting completeness before attributing the variance to demand.',
    kind: 'write-cells',
    params: { target: { range: 'Sales!E2:E3' }, cells: [['=(B2-C2)/C2'], ['=(B3-C3)/C3']] },
  },
  powerpoint: {
    title: 'Regional launch.pptx',
    body: 'Pilot results. Launch decision. Integration risks.',
    answer:
      '## Narrative review\nLead with the decision, then show the pilot evidence.\n\n1. State the requested launch decision.\n2. Show results against the success criteria.\n3. Separate confirmed risks from open questions.\n4. End with owners and decision dates.',
    kind: 'insert-slide',
    params: { text: 'Launch decision\nRecommendation: complete integration checks before launch.' },
  },
  onenote: {
    title: 'Launch research',
    body: 'Pilot findings and integration checklist.',
    answer:
      '## Evidence matrix\n\n| Question | Evidence | Still unknown |\n| --- | --- | --- |\n| Is the pilot complete? | Scope delivered | Acceptance owner |\n| Is launch ready? | Two checks remain | Completion date |\n\nKeep completion of the pilot separate from approval to launch.',
    kind: 'add-outline',
    params: {
      html: '<h2>Launch research</h2><p>Pilot scope is complete. Two integration checks remain open.</p>',
    },
  },
  outlook: {
    title: 'Re: Launch readiness',
    body: 'Can we confirm the launch date? Two checks remain open.',
    answer:
      '## Commitments\n\n| Action | Owner | Due date | Status |\n| --- | --- | --- | --- |\n| Complete integration checks | Not stated | Not stated | Open |\n| Confirm launch date | Not stated | Not stated | Requested |\n\nThe request for a date is not yet an accepted commitment.',
    kind: 'reply-mail',
    params: {
      text: 'The pilot is complete. We can confirm the launch date after the two integration checks close. Please confirm owners and target dates.',
    },
  },
  teams: {
    title: 'Launch review transcript',
    body: 'Pilot complete. Launch date proposed; no decision yet.',
    answer:
      '## Meeting recap\n- **Agreed:** the pilot delivered its scope.\n- **Proposed:** a launch next month.\n- **Unresolved:** owners and dates for the integration checks.\n\nNo launch approval was recorded in the available transcript.',
    kind: 'post-message',
    params: {
      text: 'Pilot scope is complete. Launch approval remains open pending integration checks; owners and dates are still to be assigned.',
    },
  },
};

/** A real controller over an explicitly scripted session. No credentials, network, or Office writes. */
export function makeDemoController(surface: Surface): PanelController {
  const sample = SAMPLE[surface];
  let sequence = 0;
  const refs: ContextRef[] = [
    {
      id: surface === 'excel' ? 'xl:Sales!A1:D12' : `${surface}:selection`,
      surface,
      kind: surface === 'excel' ? 'range' : 'selection',
      title: sample.title,
      preview: sample.body,
      live: true,
    },
    {
      id: 'demo:source',
      surface,
      kind: 'file',
      title: 'Pilot findings',
      preview: 'Sample source: the pilot delivered its scope. Two integration checks remain.',
    },
  ];
  const session: AssistLike = {
    context: { size: 0 },
    attachRef: async () => undefined,
    detach: () => undefined,
    async *ask() {
      yield { type: 'token', text: sample.answer };
      yield { type: 'citation', source: { title: sample.title, excerpt: sample.body } };
      yield { type: 'done' };
    },
    async *runCommands(_task: string, options?: RunCommandsOptions) {
      const request: ActuationRequest = {
        changeId: asChangeId(`demo-${++sequence}`),
        surface,
        kind: sample.kind,
        params: sample.params,
      };
      const effects = [
        {
          request,
          command: renderCommandLine(request),
          approvalClass: approvalClassOf(request.kind),
          reversible: isReversibleKind(request.kind),
          dryRun: {
            target: sample.params.target?.range ?? sample.params.target?.matchText ?? sample.title,
            resolved: sample.params.text ?? sample.params.html ?? 'Variance formulas',
          },
        },
      ];
      yield { type: 'token', text: 'This is a scripted preview. Review the sample change below.' };
      const approved = await options?.approvePlan?.(effects);
      if (options?.signal?.aborted) return;
      yield {
        type: 'token',
        text: approved
          ? '\n\nSample change approved. No Office document was modified.'
          : '\n\nSample change rejected. Nothing was applied.',
      };
      yield { type: 'done' };
    },
    async plan(task) {
      const intent =
        commandPaletteFor(surface).verbs.find((verb) =>
          ['rewrite', 'draft', 'notes'].includes(verb.intent),
        )?.intent ?? 'ask';
      return {
        plan: {
          intent,
          surface,
          scope: { kind: 'document' },
          ground: [{ kind: 'this' }],
          context: [],
          steps: ['Read the sample context', 'Stage a sample change for review'],
          excludes: [],
          clarify: [],
          raw: task,
        },
        errors: [],
        needsClarification: false,
      };
    },
    apply: async (kind, _params, changeId) => ({
      ok: true,
      kind,
      changeId,
      location: 'Scripted preview only',
    }),
    ingest: async () => undefined,
  };
  const controller = new PanelController(session, {
    listContext: async () => refs,
    canRevealContext: () => false,
  });
  controller.setDiscoveredCatalog(
    [],
    [
      {
        id: 'sample-policies',
        resourceName:
          'projects/demo/locations/global/collections/default_collection/dataStores/sample-policies',
        displayName: 'Policy library (sample)',
        connector: 'SharePoint',
      },
      {
        id: 'sample-research',
        resourceName:
          'projects/demo/locations/global/collections/default_collection/dataStores/sample-research',
        displayName: 'Research notes (sample)',
        connector: 'OneDrive',
      },
    ],
  );
  return controller;
}
