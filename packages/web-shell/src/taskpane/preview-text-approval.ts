import {
  asChangeId,
  approvalClassOf,
  isReversibleKind,
  type ActuationRequest,
} from '@ge/contracts';
import type { PanelState } from '../controller.js';

/** Approval display fixtures only. No request is dispatched by the preview controller. */
export function textApprovalPreview(params: URLSearchParams): Partial<PanelState> {
  const mode = params.get('position-approval');
  if (!['body-end', 'anchor-start', 'control', 'shape'].includes(mode ?? '')) return {};
  const pin = 'sha256:' + 'a'.repeat(64);
  const request: ActuationRequest =
    mode === 'shape'
      ? {
          changeId: asChangeId('preview-shape'),
          surface: 'powerpoint',
          kind: 'set-shape-text',
          params: {
            target: {
              slideId: 'slide-1',
              shapeId: 'shape-2',
              expectedDocumentId: pin,
              expectedTextHash: pin,
            },
            text: 'Reviewed metrics — sample only.',
          },
        }
      : mode === 'control'
        ? {
            changeId: asChangeId('preview-control'),
            surface: 'word',
            kind: 'fill-content-control',
            params: {
              target: { contentControlId: '17', expectedDocumentId: pin, expectedTextHash: pin },
              text: 'Reviewed close memo — sample only.',
            },
          }
        : {
            changeId: asChangeId('preview-position'),
            surface: 'word',
            kind: 'insert-text',
            params: {
              position: mode === 'body-end' ? 'end' : 'start',
              text: 'Reviewed text — sample only.',
              ...(mode === 'anchor-start'
                ? { target: { matchText: 'Reporting clause', contextHint: 'Section 7' } }
                : {}),
            },
          };
  return {
    messages: [
      {
        id: 'preview-approval-note',
        role: 'assistant',
        text: 'Sample approval preview. No Office operation runs here.',
      },
    ],
    chips: [],
    suggestions: [],
    proposals: [],
    steps: [],
    pendingWrite: undefined,
    error: undefined,
    busy: true,
    pendingPlan: {
      summary: '1 native text change',
      effects: [
        {
          request,
          command: '',
          approvalClass: approvalClassOf(request.kind),
          reversible: isReversibleKind(request.kind),
        },
      ],
    },
  };
}
