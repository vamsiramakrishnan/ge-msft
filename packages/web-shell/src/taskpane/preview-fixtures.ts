import type { Surface } from '@ge/contracts';
import { asChangeId, approvalClassOf, isReversibleKind } from '@ge/contracts';
import type { ProvenancePayload } from '@ge/contracts';
import type {
  ChatMessage,
  ContextChip,
  EffectDryRun,
  PanelState,
  PendingPlan,
  PendingWrite,
  PlanEffect,
  Proposal,
  RunStep,
  Skill,
  Suggestion,
} from '../controller.js';

/**
 * Realistic, hand-authored `PanelState` fixtures that exercise EVERY card the panel can render — a
 * streamed assistant message with citations, attached + available context chips, suggestions, a
 * command-loop transcript, a pending plan (multi-effect set), a pending per-write, a proposal in
 * each status, an error, and the busy flag. The preview harness composes these into a fake
 * `PanelController` so the whole UI is inspectable in a plain browser with no Office host; the
 * render smoke test reuses them to assert every card mounts. These are view fixtures only — they
 * carry no controller logic and never actuate anything.
 */

export const FIXTURE_CHIPS: ContextChip[] = [
  {
    id: 'notebook:vendor-risk',
    title: 'Notebook: Vendor Risk',
    kind: 'indexed-document',
    attached: true,
    preview: 'NotebookLM · 14 sources on vendor risk posture',
  },
  {
    id: 'sp:contracts',
    title: 'SharePoint · Contracts',
    kind: 'file',
    attached: true,
    preview: 'Sites.Selected · /sites/Procurement/Contracts',
  },
  {
    id: 'word:selection',
    title: 'Selection',
    kind: 'selection',
    attached: false,
    preview: '“…the liability cap excludes data-breach events…”',
  },
  {
    id: 'word:document',
    title: 'Whole document',
    kind: 'document',
    attached: false,
    preview: 'Vendor_Risk_Memo.docx · 8 pages',
  },
];

export const FIXTURE_MESSAGES: ChatMessage[] = [
  {
    id: 'u-1',
    role: 'user',
    text: 'Summarise the FY26 risk exposure for Northwind Cloud and flag anything below policy.',
  },
  {
    id: 'a-1',
    role: 'assistant',
    text: 'Northwind Cloud is a Tier-1 vendor at $480k/yr. Their SLA of 99.5% sits below your 99.9% policy, and the liability cap has no data-breach carve-out — both are flagged. The contract renews in November 2026.',
    sources: [
      {
        title: 'Vendor Risk (Notebook)',
        uri: 'https://example.com/nb/vendor-risk',
        locator: '§3.2',
      },
      {
        title: 'MSA — Northwind Cloud',
        uri: 'https://example.com/contracts/northwind',
        locator: 'cl. 11',
      },
      { title: 'SLA addendum', locator: 'p. 4' },
    ],
  },
  {
    id: 'a-2',
    role: 'assistant',
    text: 'Comparing the renewal terms against your data-residency policy now',
    streaming: true,
  },
];

export const FIXTURE_SUGGESTIONS: Suggestion[] = [
  {
    id: 's-1',
    title: 'Draft a renewal-risk note for this vendor',
    detail: 'Grounded on the unit · lands as a tracked change',
    query: 'Draft a short renewal-risk note for Northwind Cloud.',
  },
];

export const FIXTURE_STEPS: RunStep[] = [
  { id: 'step-1', kind: 'turn-start', text: 'Turn 1' },
  { id: 'step-2', kind: 'command', text: 'read Sales!A2:C7' },
  { id: 'step-3', kind: 'read-result', text: 'read vendor rows' },
  { id: 'step-4', kind: 'command', text: 'set Sales!F2 =C2-D2' },
  { id: 'step-5', kind: 'write-result', text: 'write-cells — applied' },
  { id: 'step-6', kind: 'done', text: 'Done' },
];

const planEffect = (
  id: string,
  kind: PlanEffect['request']['kind'],
  command: string,
  params: PlanEffect['request']['params'],
  dryRun?: EffectDryRun,
): PlanEffect => ({
  command,
  request: { changeId: asChangeId(id), kind, surface: 'excel', params },
  approvalClass: approvalClassOf(kind),
  reversible: isReversibleKind(kind),
  ...(dryRun ? { dryRun } : {}),
});

export const FIXTURE_PLAN: PendingPlan = {
  summary: '3 writes + 2 comments',
  effects: [
    planEffect(
      'p-1',
      'write-cells',
      'set Sales!F2 =C2-D2',
      { target: { range: 'Sales!F2' }, cells: [['=C2-D2']] },
      { target: 'Sales!F2', resolved: '=C2-D2', before: '(empty)', after: '$184,000' },
    ),
    planEffect(
      'p-2',
      'write-cells',
      'set Sales!F3 =C3-D3',
      { target: { range: 'Sales!F3' }, cells: [['=C3-D3']] },
      { target: 'Sales!F3', resolved: '=C3-D3', before: '(empty)', after: '$71,500' },
    ),
    planEffect(
      'p-3',
      'write-cells',
      'set Sales!F4 =C4-D4',
      { target: { range: 'Sales!F4' }, cells: [['=C4-D4']] },
      { target: 'Sales!F4', resolved: '=C4-D4', before: '(empty)', after: '$22,800' },
    ),
    planEffect(
      'p-4',
      'add-comment',
      'comment Sales!F2 "Below margin policy — review"',
      { target: { range: 'Sales!F2' }, text: 'Below margin policy — review' },
      { target: 'Sales!F2', resolved: 'Below margin policy — review' },
    ),
    planEffect('p-5', 'add-comment', 'comment Sales!F4 "Confirm residency clause"', {
      target: { range: 'Sales!F4' },
      text: 'Confirm residency clause',
    }),
  ],
};

export const FIXTURE_PENDING_WRITE: PendingWrite = {
  changeId: asChangeId('w-1'),
  kind: 'tracked-change',
  command: 'suggest "SLA of 99.5%" => "SLA of 99.5% (below the 99.9% policy)"',
};

/** Provenance stamped onto an applied write — drives the proposal drill-down fixture. */
export const FIXTURE_PROVENANCE: ProvenancePayload = {
  agentId: 'contract-review-agent@v2',
  identity: 'v.k@acme.com',
  timestamp: '2026-06-23T09:14:02.000Z',
  sources: [
    { title: 'Vendor Risk Policy v4', uri: 'https://example.com/policy/v4', locator: '§3.2' },
    {
      title: 'MSA — Northwind Cloud',
      uri: 'https://example.com/contracts/northwind',
      locator: 'cl. 11',
    },
  ],
  contentHash: 'sha256:4f1a9c2e7b08d3c6a51e0f9b2d7c4a8e3f6b1d0c9a2e5b8f',
};

export const FIXTURE_PROPOSALS: Proposal[] = [
  {
    changeId: asChangeId('c-1'),
    kind: 'tracked-change',
    params: {
      target: { matchText: 'SLA of 99.5%' },
      text: 'SLA of 99.9% (FSI standard)',
    },
    label: 'Raise the SLA to policy as a tracked change',
    status: 'pending',
  },
  {
    changeId: asChangeId('c-2'),
    kind: 'write-cells',
    params: { target: { range: 'Sales!F2' }, cells: [['=C2-D2']] },
    label: 'Write the net exposure formula to F2',
    status: 'pending',
    entityCard: {
      title: 'Northwind Cloud',
      subtitle: 'Entity · enriched by Gemini Enterprise',
      rows: [
        { key: 'Contract status', value: 'Active · renews Nov 26' },
        { key: 'SLA', value: '99.5% (below policy)' },
        { key: 'ISO 27001', value: 'Valid' },
        { key: 'Open risks', value: '2 (liability cap)' },
      ],
      footnote: 'Loaded from the unit · not stored in the workbook',
    },
  },
  {
    changeId: asChangeId('c-3'),
    kind: 'add-comment',
    params: { target: { matchText: 'renews' }, text: 'Renewal window opens 90 days prior.' },
    label: 'Add a renewal-window comment',
    status: 'applied',
    detail: 'Landed at cl. 11 · reversible',
    provenance: FIXTURE_PROVENANCE,
  },
  {
    changeId: asChangeId('c-4'),
    kind: 'write-cells',
    params: { target: { range: 'Sales!G2' }, cells: [['=F2*1.1']] },
    label: 'Write the adjusted exposure to G2',
    status: 'degraded',
    detail: 'Anchor drifted — surfaced as a panel item instead.',
  },
];

export const FIXTURE_SKILLS: Skill[] = [
  {
    name: 'flag-vendor-risk',
    description:
      'Read the vendor row, compare SLA + liability against policy, and propose a flag plan.',
    params: [
      { name: 'vendor', example: 'Northwind Cloud' },
      { name: 'tier', example: '1' },
    ],
    registered: true,
    def: 'def flag-vendor-risk(vendor, tier) { read row; compare policy; suggest + comment }',
  },
  {
    name: 'draft-renewal-note',
    description: 'Draft a short, grounded renewal-risk note and land it as a tracked change.',
    params: [{ name: 'vendor', example: 'Northwind Cloud' }],
    registered: true,
    def: 'def draft-renewal-note(vendor) { ground unit; suggest tracked-change }',
  },
];

export const FIXTURE_ERROR = 'Could not reach the grounding endpoint — retry in a moment.';

/** A full, everything-on snapshot used as the preview's default and by the smoke test. */
export const FIXTURE_STATE: PanelState = {
  messages: FIXTURE_MESSAGES,
  chips: FIXTURE_CHIPS,
  suggestions: FIXTURE_SUGGESTIONS,
  proposals: FIXTURE_PROPOSALS,
  changes: [],
  steps: FIXTURE_STEPS,
  pendingWrite: FIXTURE_PENDING_WRITE,
  pendingPlan: FIXTURE_PLAN,
  skills: FIXTURE_SKILLS,
  busy: true,
  error: FIXTURE_ERROR,
};

export const PREVIEW_SURFACES: Surface[] = ['word', 'excel', 'powerpoint', 'outlook', 'teams'];
