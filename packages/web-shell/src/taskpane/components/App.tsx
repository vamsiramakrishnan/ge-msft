import { AnalysisWorkbench } from './AnalysisWorkbench.js';
import { WorkflowWorkbench } from './WorkflowWorkbench.js';
import { EvidencePanel } from './EvidencePanel.js';
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  deriveOutput,
  actionParameters,
  hasUnfilledPlaceholder,
  GROUND_SOURCE_TO_SELECTION_KIND,
  type Surface,
  type ChangeId,
  type QuickAction,
  type Intent,
  type GroundingSelection,
} from '@ge/contracts';
import {
  resolveGrounding,
  type GroundingResolveContext,
  type ResolvedGrounding,
} from '@ge/gemini-client';
import type { ContextChip, PanelController } from '../../controller.js';
import { usePanelState } from '../usePanelState.js';
import { SavedDiagnostics, browserStorage } from '../../saved-diagnostics.js';
import { ResumeStore } from '../../resume-store.js';
import { detectLanguage, uiStrings } from '../../ui-language.js';
import { MessageThread } from './MessageThread.js';
import { Composer, type ComposerInvocation, type ComposerMention } from './Composer.js';
import { invocationToSeed, quickActionToInvocation } from './quick-action-seed.js';
import { QuickActionParamForm } from './QuickActionParamForm.js';
import { ProposalCard } from './ProposalCard.js';
import { RunSteps } from './RunSteps.js';
import { WriteApprovalCard } from './WriteApprovalCard.js';
import { PlanApprovalCard } from './PlanApprovalCard.js';
import { ShareApprovalCard } from './ShareApprovalCard.js';
import { CommandPlanCard } from './CommandPlanCard.js';
import { PlanClarificationCard } from './PlanClarificationCard.js';
import { AnswerOptions, type AnswerOptionsValue } from './AnswerOptions.js';
import { surfacePrimaryActions } from './SurfaceCommandCenter.js';
import { ContextStrip } from './ContextStrip.js';
import { WorkspaceHome } from './WorkspaceHome.js';
import { Toolbar } from './Toolbar.js';
import { appItem, appName } from './app-names.js';
import { extractDirectCommandProgram } from '../direct-command.js';
import { buildInsertArtifactProgram, type InsertableArtifact } from '../insert-artifact.js';

export interface AppProps {
  controller: PanelController;
  surface: Surface;
  agentLabel?: string;
  /**
   * The intents the active surface can actually run (from `intentsForManifest(bridge manifest)`).
   * Narrows the quick-action bar and the `/` palette to capability closure (ADR-0006). When omitted
   * (preview/tests), only the per-surface filter applies.
   */
  allowedIntents?: Iterable<Intent>;
  /** The signed-in account's hash for saved diagnostics; without it nothing is saved. */
  diagnosticsScope?: string;
  /** The signed-in user's name for the home greeting; the greeting has no name without it. */
  userName?: string;
  /** Build-time model override; blank leaves model choice to the engine. */
  initialModelId?: string;
  onModelChange?: (modelId: string | undefined) => void;
}

/** Commit and build time in every Copy diagnostics, so a pasted report names its build. */
const BUILD = typeof __BUILD_ID__ === 'string' && __BUILD_ID__ ? { build: __BUILD_ID__ } : {};

const SURFACE_PLACEHOLDER: Readonly<Record<string, string>> = {
  word: 'Ask about the selection…',
  excel: 'Ask about this range…',
  powerpoint: 'Ask about this slide…',
  onenote: 'Ask about this page…',
  outlook: 'Ask about this email…',
  teams: 'Ask about this meeting…',
};

/**
 * THE single routing predicate (EXPERIENCE.md §3) — total over the 7-verb set. Chat verbs
 * (`ask`/`summarize`/`explain`) and a bare question (undefined intent) are single-shot reads → `send`;
 * the four specialist verbs (`rewrite`/`review`/`draft`/`notes`) land an effect through the gate →
 * `runCommands`. A `rewrite` can NEVER reach `send`. App makes this decision ONCE, here, used by both
 * the quick-action chips and the composer submit.
 */
export function isActuating(intent: Intent | undefined): boolean {
  // Single source of truth (no second hand-maintained verb set to drift): a verb actuates iff its
  // closure-derived output is not `chat`. `deriveOutput` lives in contracts next to INTENT_REQUIRES,
  // so the routing decision and the capability closure can never disagree (security review, Finding 2).
  return intent !== undefined && deriveOutput(intent) !== 'chat';
}

/** Words/markers that signal a multi-step or constrained instruction worth a plan-confirm first. */
const CONSTRAINT_MARKERS =
  /(^|\s)(but|only|except|without|unless|then|leave|keep|preserve|ignore|excluding)(\s|$)|[,;]/i;

/**
 * The heuristic for "complex free-text" (EXPERIENCE.md §F): an instruction earns the planner-confirm
 * front door when it's long (≥ 12 words) or carries a constraint/exclusion marker (",", "but",
 * "only", "leave …", etc.). A short, single-shot instruction ("make it formal") skips the planner and
 * goes straight to the executor. Conservative by design — the executor's gate is the backstop either way.
 */
export function isComplexInstruction(instruction: string): boolean {
  const t = instruction.trim();
  if (!t) return false;
  return t.split(/\s+/).length >= 12 || CONSTRAINT_MARKERS.test(t);
}

/**
 * The polite / first-person lead-in a request may open with before its action verb: "please",
 * "hi, can you", "I want to", "I'd like you to", "we need to", "let's", "help me", in any order.
 * Each phrase must end on a word boundary, so "I want to know how to add…" never reaches a verb.
 */
const LEAD_IN =
  String.raw`(?:(?:ok(?:ay)?|yes|yeah|yep|sure|alright|great|cool|thanks|hey|hi|hello|so|now|also|and|then` +
  String.raw`|please|pls|plz|kindly|just|quickly|maybe|go\s+ahead\s+and` +
  String.raw`|(?:can|could|would|will)\s+(?:you|u)` +
  String.raw`|(?:i|we)\s*(?:would|['’]d)\s+(?:like|love)\s+(?:(?:you|u)\s+)?to` +
  String.raw`|(?:i|we)\s+(?:want|need|have|plan|wish|intend)\s+(?:(?:you|u)\s+)?to` +
  String.raw`|(?:i|we)\s+(?:wanna|gotta)` +
  String.raw`|(?:i\s*['’]m|i\s+am|we\s*['’]re|we\s+are)\s+(?:trying|going|looking|hoping)\s+to` +
  String.raw`|let\s*['’]?s|let\s+me|help\s+me(?:\s+to)?|try\s+to` +
  // Indonesian: "tolong", "mohon", "bisakah kamu", "saya mau/ingin", "bantu saya".
  String.raw`|tolong|mohon|coba|silah?kan|ayo|bantu(?:\s+saya)?` +
  String.raw`|(?:bisa|bisakah|dapatkah|boleh|bolehkah)(?:\s+(?:kamu|anda|kau))?` +
  String.raw`|(?:saya|aku|kami|kita)\s+(?:mau|ingin|pengen|perlu|harus|akan)` +
  String.raw`)(?:[\s,!.]+|$))*`;

/** A request whose (lead-in-stripped) first word is one of `verbs`, optionally followed by `tail`. */
function requestRe(verbs: string, tail = ''): RegExp {
  return new RegExp(String.raw`^\s*` + LEAD_IN + String.raw`(?:${verbs})\b` + tail, 'i');
}

/** Indonesian verbs that CREATE something (make/add/insert/compose) — not delete or change. */
const ID_CREATE_VERBS =
  'buat|buatkan|membuat|bikin|bikinkan|jadikan|tambah|tambahkan|menambah|menambahkan|' +
  'sisipkan|menyisipkan|masukkan|memasukkan|susun|menyusun|draf|drafkan';

const EXCEL_CHART_CREATE_RE = requestRe(
  'create|make|insert|add|build|generate|plot|visuali[sz]e|' + ID_CREATE_VERBS,
  String.raw`[\s\S]*\b(?:chart|graph|visuali[sz]ation|grafik|diagram|bagan)\b`,
);
const EXCEL_CHART_CONVERT_RE = requestRe(
  'turn|convert',
  String.raw`[\s\S]*\b(?:into|to)\b[\s\S]*\b(?:chart|graph|visuali[sz]ation)\b`,
);
const WORD_REWRITE_RE = requestRe(
  'rewrite|revise|tighten|edit|replace|improve',
  String.raw`[\s\S]*\b(?:selection|selected text|paragraph|text|wording|clause|sentence)\b`,
);
/** Indonesian: "tulis ulang / perbaiki / revisi … teks|paragraf|kalimat|klausul|yang dipilih". */
const WORD_REWRITE_ID_RE = requestRe(
  'tulis\\s+ulang|tuliskan\\s+ulang|perbaiki|revisi|sunting|perhalus|perjelas|persingkat|rapikan',
  String.raw`[\s\S]*\b(?:teks|tulisan|paragraf|kalimat|klausul|kata-kata|pilihan|bagian|dipilih)\b`,
);
const WORD_REVIEW_RE = requestRe(
  'review|comment|flag|mark',
  String.raw`[\s\S]*\b(?:issue|issues|risk|risks|gap|gaps|claim|claims|comment|comments)\b`,
);
/** Indonesian: "tinjau / tandai / komentari … klaim|risiko|masalah|celah|komentar". */
const WORD_REVIEW_ID_RE = requestRe(
  'tinjau|tandai|komentari|beri\\s+komentar',
  String.raw`[\s\S]*\b(?:klaim|risiko|masalah|celah|isu|komentar)\b`,
);
const POWERPOINT_DRAFT_RE = requestRe(
  'create|make|insert|add|build|generate|draft|' + ID_CREATE_VERBS,
  String.raw`[\s\S]*\b(?:slides?|salindia)\b`,
);
const OUTLOOK_DRAFT_RE = requestRe(
  'draft|write|compose|create|buat|buatkan|membuat|tulis|tuliskan|menulis|susun|menyusun',
  String.raw`[\s\S]*\b(?:reply|email|mail|message|balasan|surel|pesan)\b`,
);

function intentAllowed(intent: Intent, allowedIntents: Iterable<Intent> | undefined): boolean {
  if (allowedIntents === undefined) return true;
  for (const allowed of allowedIntents) {
    if (allowed === intent) return true;
  }
  return false;
}

/**
 * Narrow natural-language promotion for bridge-backed mutations. Plain "create a chart..." / "draft
 * a slide..." requests are actuating requests, but historically landed in read-only chat unless the
 * user knew to type a slash verb. Promote only imperative language, only when the runtime capability
 * closure includes the matching intent. The promoted turn still enters `runCommands`, so it dry-runs,
 * previews, and waits for explicit approval before any host mutation.
 */
export function inferImplicitIntent(
  surface: Surface,
  allowedIntents: Iterable<Intent> | undefined,
  inv: ComposerInvocation,
): Intent | undefined {
  if (inv.intent !== undefined) return inv.intent;
  const raw = inv.raw.trim();
  if (!raw || raw.startsWith('/')) return undefined;
  switch (surface) {
    case 'excel':
      if (
        intentAllowed('visualize', allowedIntents) &&
        (EXCEL_CHART_CREATE_RE.test(raw) || EXCEL_CHART_CONVERT_RE.test(raw))
      ) {
        return 'visualize';
      }
      return undefined;
    case 'word':
      if (
        intentAllowed('rewrite', allowedIntents) &&
        (WORD_REWRITE_RE.test(raw) || WORD_REWRITE_ID_RE.test(raw))
      )
        return 'rewrite';
      if (
        intentAllowed('review', allowedIntents) &&
        (WORD_REVIEW_RE.test(raw) || WORD_REVIEW_ID_RE.test(raw))
      )
        return 'review';
      return undefined;
    case 'powerpoint':
      return intentAllowed('draft', allowedIntents) && POWERPOINT_DRAFT_RE.test(raw)
        ? 'draft'
        : undefined;
    case 'outlook':
      return intentAllowed('draft', allowedIntents) && OUTLOOK_DRAFT_RE.test(raw)
        ? 'draft'
        : undefined;
    case 'onenote':
    case 'teams':
      return undefined;
  }
}

function hasActuatingIntent(allowedIntents: Iterable<Intent> | undefined): boolean {
  if (allowedIntents === undefined) return true;
  for (const intent of allowedIntents) {
    if (isActuating(intent)) return true;
  }
  return false;
}

/**
 * A leading `/word` that is not one of this surface's verbs carries no intent (the composer resolves
 * verbs against the surface palette only), so route the line, without its slash, as free text. Users type
 * command names they have seen: live 2026-10-09, "/shape Add a blue rectangle to the top right of
 * slide 2" skipped the planner, reached chat, which cannot write, and was told shapes are not
 * supported. A bare `/word` with nothing after it is left alone.
 */
export function withoutUnknownVerb(inv: ComposerInvocation): ComposerInvocation {
  const raw = inv.raw.trim();
  if (inv.intent !== undefined || !raw.startsWith('/') || !inv.instruction.trim()) return inv;
  // Only the slash goes: the word can be the user's verb ("/delete slide 3").
  const text = raw.slice(1);
  return { ...inv, raw: text, instruction: text };
}

/**
 * Free text on a writable surface goes to the planner. Its structured `changes` signal chooses
 * chat or execution; guessing from question words would discard mixed question-and-write requests.
 * Explicit intents and pasted CLI retain their dedicated routes.
 */
export function shouldUsePlannerForFreeText(
  allowedIntents: Iterable<Intent> | undefined,
  inv: ComposerInvocation,
): boolean {
  if (inv.intent !== undefined) return false;
  const raw = inv.raw.trim();
  if (!raw || raw.startsWith('/')) return false;
  if (!hasActuatingIntent(allowedIntents)) return false;
  return true;
}

/**
 * Map ONE typed composer `@`-mention to its {@link GroundingSelection} (Finding #2/#B-wire). The
 * `GroundSource` kind picks the value-level selection kind; the reference-kinds (`current-context`,
 * `unit`) carry no id, while `document`/`person`/`data-store`/`upload` need the addressable handle the
 * mention carried (`ref`). A mention of an addressable kind with NO ref cannot be resolved — it is
 * dropped (returns `undefined`) rather than smuggled into the prompt as raw text.
 */
export function mentionToSelection(m: ComposerMention): GroundingSelection | undefined {
  const kind = GROUND_SOURCE_TO_SELECTION_KIND[m.kind];
  switch (kind) {
    case 'current-context':
    case 'unit':
      return { kind };
    case 'document':
    case 'person':
    case 'data-store':
      return m.ref ? { kind, id: m.ref } : undefined;
    case 'upload':
      return m.ref ? { kind, fileId: m.ref } : undefined;
    default:
      return undefined;
  }
}

/**
 * Turn an invocation's typed `@`-mentions into STRUCTURED grounding (Finding #2/#B-wire): map each
 * mention to a {@link GroundingSelection}, then resolve them onto the streamAssist request buckets via
 * `resolveGrounding` (gemini-client). The mentions become addressed `queryParts`/`dataStoreSpecs`/
 * `fileIds` — NEVER inlined into the prompt string; the raw text is kept only for audit/display.
 * Returns `undefined` when no mention resolved to anything, so a mention-free turn carries no grounding.
 */
export function invocationToGrounding(
  inv: ComposerInvocation,
  ctx: GroundingResolveContext = {},
): ResolvedGrounding | undefined {
  const selections = inv.mentions
    .map(mentionToSelection)
    .filter((s): s is GroundingSelection => s !== undefined);
  if (selections.length === 0) return undefined;
  return resolveGrounding(selections, ctx);
}

function excelSelectionRange(chips: ContextChip[]): string | undefined {
  const selection = chips.find(
    (chip) =>
      chip.kind === 'range' && chip.id.startsWith('xl:') && !chip.id.startsWith('xl:named:'),
  );
  return selection?.title;
}

/**
 * The task pane. A thin React view over `PanelController` state: header (agent identity), context
 * tray (attach/detach chips), streamed grounded thread with citations, proposal-review cards, and
 * the composer (send / cancel). No host or network code here — the controller owns all of that.
 */
export function App({
  controller,
  surface,
  agentLabel,
  allowedIntents,
  diagnosticsScope,
  userName,
  initialModelId,
  onModelChange,
}: AppProps): JSX.Element {
  const state = usePanelState(controller);
  const savedStore = useMemo(
    () => new SavedDiagnostics(browserStorage(), diagnosticsScope),
    [diagnosticsScope],
  );
  const [savedDiagnostics, setSavedDiagnostics] = useState(() => savedStore.list());
  // A change task the pane was closed during is offered for finishing (per signed-in user).
  useEffect(() => {
    void controller.attachResumeStore(new ResumeStore(browserStorage(), diagnosticsScope));
  }, [controller, diagnosticsScope]);
  // The question saved with diagnostics is what the user TYPED (or the quick action they chose),
  // captured when they sent it. Never read back from the thread: a bubble can hold an inserted
  // answer or history loaded from the server, which must not reach browser storage.
  const typedQuestion = useRef<{ id: string; text: string }>();
  const paneLoad = useMemo(() => Date.now().toString(36), []);
  const questionSeq = useRef(0);
  const recordQuestion = (text: string): void => {
    const t = text.trim();
    if (t) typedQuestion.current = { id: `${paneLoad}:${++questionSeq.current}`, text: t };
  };
  // Activity's Copy diagnostics carries the question too, in the same shape as a saved entry.
  const activityDiagnostics = (): string => {
    const diagnostics = controller.diagnostics({ surface, ...BUILD });
    const question = typedQuestion.current?.text;
    if (!question) return diagnostics;
    return JSON.stringify({ question, diagnostic: JSON.parse(diagnostics) as unknown }, null, 2);
  };
  const wasBusy = useRef(false);
  useEffect(() => {
    const question = typedQuestion.current;
    if (wasBusy.current && !state.busy && question) {
      try {
        savedStore.save({
          id: question.id,
          question: question.text,
          diagnostic: JSON.parse(controller.diagnostics({ surface, ...BUILD })) as unknown,
        });
        setSavedDiagnostics(savedStore.list());
      } catch {
        /* Diagnostics are best effort; never break the pane over them. */
      }
    }
    wasBusy.current = state.busy;
    // eslint-disable-next-line react-hooks/exhaustive-deps -- activityDiagnostics reads refs only
  }, [state.busy, controller, surface, savedStore]);
  // The parameterized action awaiting its `{{name}}` fill values (Workstream H), or undefined.
  const [paramFill, setParamFill] = useState<QuickAction | undefined>(undefined);
  // Answer format and style live in Settings; a new surface starts from the defaults.
  const [answerOptions, setAnswerOptions] = useState<AnswerOptionsValue>({});
  const [modelId, setModelId] = useState(initialModelId ?? '');
  useEffect(() => setAnswerOptions({}), [surface]);
  const hasBlockingGate = Boolean(
    state.pendingCommandPlan ?? state.pendingPlan ?? state.pendingWrite ?? state.pendingShare,
  );
  const actionBlocked = state.busy || hasBlockingGate;

  // Keep the newest message in view: follow the thread while the user is at the bottom, stop when
  // they scroll up to read, and jump back down when they send a new message.
  const threadRef = useRef<HTMLElement>(null);
  const followThread = useRef(true);
  const seenMessages = useRef(0);
  useLayoutEffect(() => {
    const region = threadRef.current;
    const count = state.messages.length;
    if (count < seenMessages.current) followThread.current = true; // a new chat
    if (state.messages.slice(seenMessages.current).some((m) => m.role === 'user')) {
      followThread.current = true;
    }
    seenMessages.current = count;
    if (region && followThread.current) region.scrollTop = region.scrollHeight;
  }, [state.messages, state.steps, state.proposals, state.error, state.pendingPlanClarification]);
  const onThreadScroll = (): void => {
    const region = threadRef.current;
    if (!region) return;
    followThread.current = region.scrollHeight - region.scrollTop - region.clientHeight < 48;
  };

  const primaryActions = useMemo(
    () => surfacePrimaryActions(surface, allowedIntents),
    [allowedIntents, surface],
  );
  const primaryActionIds = useMemo(
    () => primaryActions.map((action) => action.id),
    [primaryActions],
  );
  const proposalCount = state.proposals.filter((proposal) => proposal.status === 'pending').length;
  const insertExcelRange = surface === 'excel' ? excelSelectionRange(state.chips) : undefined;
  const insertArtifactDisabledReason = actionBlocked
    ? 'Finish the current turn or approval first.'
    : surface === 'excel' && !insertExcelRange
      ? 'Select a destination range first.'
      : undefined;

  // Load the attachable-context chips once on mount.
  useEffect(() => {
    void controller.refreshContext();
  }, [controller]);

  const onToggle = (id: string, attach: boolean): void => {
    if (actionBlocked) return;
    if (attach) void controller.attach(id);
    else controller.detach(id);
  };

  // Dispatch the typed invocation through ONE predicate-routed seam (EXPERIENCE.md §3): chat verbs
  // route to `send`, write/annotation verbs to the fail-closed plan/approval loop (`runCommands`).
  // The model-facing TASK stays the CLI seed (deterministic from the typed fields); ALONGSIDE it, the
  // typed `@`-mentions are CONSUMED into STRUCTURED grounding (Finding #2/#B-wire) via
  // `invocationToGrounding` → `resolveGrounding`, and passed as the turn's grounding — never discarded
  // nor forwarded only as raw text. No new gate is introduced; grounding only scopes the existing route.
  const dispatch = (typed: ComposerInvocation): void => {
    if (actionBlocked) return;
    recordQuestion(typed.raw);
    const clarificationAnswer = typed.raw.trim();
    if (
      state.pendingPlanClarification &&
      clarificationAnswer &&
      typed.intent === undefined &&
      !clarificationAnswer.startsWith('/')
    ) {
      void controller.answerPlanClarification(clarificationAnswer);
      return;
    }
    const directProgram = extractDirectCommandProgram(typed.raw);
    if (directProgram) {
      void controller.runDirectCommands(directProgram);
      return;
    }
    const inv = withoutUnknownVerb(typed);
    const effectiveIntent = inferImplicitIntent(surface, allowedIntents, inv);
    const routedInv =
      effectiveIntent !== undefined && effectiveIntent !== inv.intent
        ? { ...inv, intent: effectiveIntent, raw: `/${effectiveIntent} ${inv.raw}`.trim() }
        : inv;
    const seed = invocationToSeed(routedInv);
    // Fail-closed (Workstream H): a typed `{{name}}` slot must be filled before dispatch — never send
    // a literal placeholder to the model. A parameterized chip is collected via QuickActionParamForm
    // first, so this guard only fires on a defective seed; drop it rather than actuate on raw braces.
    if (hasUnfilledPlaceholder(seed)) return;
    if (!typed.raw.trim()) recordQuestion(seed); // a quick action: its own catalog prompt
    const grounding = invocationToGrounding(inv);
    // EXPERIENCE.md §F — the planner-confirm front door. Composer-origin natural-language actions
    // always go to the command planner first, whether the user typed `/rewrite ...`, "create a
    // chart", or "okay add this". The planner, not this router, chooses rewrite/review/etc. and can
    // ask clarifying questions before the executor sees anything. Explicit pasted CLI is handled
    // above by runDirectCommands because it is already a concrete command program.
    const composerOrigin = inv.raw.trim() !== '';
    const shouldPlanFirst =
      composerOrigin &&
      (isActuating(effectiveIntent) ||
        (effectiveIntent === undefined && shouldUsePlannerForFreeText(allowedIntents, inv)));
    if (shouldPlanFirst) {
      void controller.proposePlan(seed, grounding);
    } else if (isActuating(effectiveIntent)) {
      void controller.runCommands(seed, grounding);
    } else {
      void controller.send(seed, grounding);
    }
  };

  // A chip is the same typed Invocation a composer line is — pre-filled, then dispatched through the
  // one shared `dispatch`. (`action.output` and `isActuating` agree by construction, ADR-0006.) A
  // parameterized action (Workstream H) opens the fill form FIRST instead of dispatching, so its
  // `{{name}}` slots are collected before the typed invocation is built.
  const onQuickAction = (action: QuickAction): void => {
    if (actionBlocked) return;
    if (actionParameters(action).length > 0) setParamFill(action);
    else dispatch(quickActionToInvocation(action));
  };

  // Dispatch a parameterized action once its fill values are collected: substitute the values into
  // the prompt, build the SAME typed invocation a bare chip would (scope/intent/grounding intact),
  // and route it through the one shared `dispatch`.
  const onParamFillSubmit = (action: QuickAction, values: Record<string, string>): void => {
    setParamFill(undefined);
    dispatch(quickActionToInvocation(action, values));
  };

  // A structured composer submit (`/verb` intent + scope + `@`-mentions + instruction) → the same path.
  const onInvoke = (inv: ComposerInvocation): void => dispatch(inv);

  const onInsertArtifact = (artifact: InsertableArtifact): void => {
    if (insertArtifactDisabledReason) return;
    const built = buildInsertArtifactProgram(surface, artifact, { excelRange: insertExcelRange });
    if (!built.ok) return;
    void controller.runDirectCommands(built.program);
  };

  return (
    <div className="panel" data-surface={surface} aria-busy={state.busy}>
      <Toolbar
        surface={surface}
        allowedIntents={allowedIntents}
        agentLabel={agentLabel}
        busy={actionBlocked}
        hasGate={hasBlockingGate}
        messageCount={state.messages.length}
        proposalCount={proposalCount}
        conversations={state.conversations}
        primaryActionIds={primaryActionIds}
        hasSettings
        settingsPanel={
          <>
            <label className="field">
              <span className="field-label">Model</span>
              <select
                className="field-control"
                aria-label="Model"
                value={modelId}
                disabled={actionBlocked}
                onChange={(event) => {
                  if (actionBlocked) return;
                  const next = event.target.value;
                  setModelId(next);
                  onModelChange?.(next || undefined);
                }}
              >
                <option value="">Default engine</option>
                <option value="gemini-4-argon">Gemini 4 Argon</option>
              </select>
            </label>
            <AnswerOptions
              value={answerOptions}
              disabled={actionBlocked}
              onChange={setAnswerOptions}
            />
          </>
        }
        tools={
          <>
            <WorkflowWorkbench state={state} controller={controller} disabled={actionBlocked} />
            <AnalysisWorkbench state={state} controller={controller} disabled={actionBlocked} />
          </>
        }
        onRefreshConversations={() => void controller.refreshConversations()}
        onResumeConversation={(name) => {
          typedQuestion.current = undefined;
          void controller.resumeConversation(name);
        }}
        onNewConversation={() => {
          typedQuestion.current = undefined;
          controller.newConversation();
        }}
        savedDiagnostics={savedDiagnostics}
        onClearSavedDiagnostics={() => {
          savedStore.clear();
          setSavedDiagnostics([]);
        }}
        onQuickAction={onQuickAction}
      >
        <ContextStrip
          chips={state.chips}
          disabled={actionBlocked}
          onToggle={onToggle}
          onReveal={(id) => void controller.reveal(id)}
          onRefresh={() => void controller.refreshContext()}
        />

        {state.suggestions.length > 0 && (
          <section className="suggestions" aria-label="Suggestions">
            {state.suggestions.map((s) => (
              <button
                key={s.id}
                type="button"
                className="suggestion"
                disabled={actionBlocked}
                onClick={() => {
                  if (s.query) controller.onAutomate(s.query);
                  controller.dismissSuggestion(s.id);
                }}
              >
                <span className="s-title">{s.title}</span>
                {s.detail && <span className="s-detail">{s.detail}</span>}
              </button>
            ))}
          </section>
        )}

        <main
          className="thread-region"
          aria-label="Conversation and activity"
          ref={threadRef}
          onScroll={onThreadScroll}
        >
          <EvidencePanel evidence={state.evidence} />
          {state.resume && !state.busy ? (
            <ResumeCard
              request={state.resume.request}
              made={state.resume.changes.length}
              onFinish={() => void controller.finishResume()}
              onDismiss={() => controller.dismissResume()}
            />
          ) : null}
          {state.messages.length === 0 && (
            <WorkspaceHome
              surface={surface}
              userName={userName}
              actions={primaryActions}
              disabled={actionBlocked}
              onAction={onQuickAction}
            />
          )}
          {state.conversationFrom && state.conversationFrom !== surface ? (
            <p className="conversation-from" role="note">
              This chat started in {appName(state.conversationFrom)}. Earlier answers are about that{' '}
              {appItem(state.conversationFrom)}; new questions use this {appItem(surface)}.
            </p>
          ) : null}
          <MessageThread
            messages={state.messages}
            showEmpty={false}
            disabled={actionBlocked}
            surface={surface}
            onInsertArtifact={onInsertArtifact}
            onRevealLocation={(target) => void controller.revealLocation(surface, target)}
            insertArtifactDisabledReason={insertArtifactDisabledReason}
            language={state.language ?? 'en'}
            onUndoChanges={(id) => void controller.undoChanges(id)}
            onRetry={() => void controller.retryLast()}
          />

          <PlanClarificationCard
            language={state.language ?? 'en'}
            pending={state.pendingPlanClarification}
            disabled={state.busy}
            onAnswer={(answer) => controller.answerPlanClarification(answer)}
          />

          <RunSteps steps={state.steps} diagnostics={activityDiagnostics} />

          <ProposalCard
            proposals={state.proposals}
            disabled={actionBlocked}
            onApply={(id: ChangeId) => {
              if (!actionBlocked) void controller.applyProposal(id);
            }}
          />

          {state.error && (
            <div className="panel-error" role="alert">
              {state.error}
            </div>
          )}
        </main>

        {hasBlockingGate && (
          <section className="gate-rail" aria-label="Decision required">
            <CommandPlanCard
              language={state.language ?? 'en'}
              pending={state.pendingCommandPlan}
              onConfirm={() => controller.confirmCommandPlan()}
              onChange={() => controller.changeCommandPlan()}
              onCancel={() => controller.cancelCommandPlan()}
            />

            <PlanApprovalCard
              language={state.language ?? 'en'}
              plan={state.pendingPlan}
              onRevealTarget={(target) => void controller.revealLocation(surface, target)}
              onApprove={() => controller.approvePlan()}
              onReject={() => controller.rejectPlan()}
            />

            <WriteApprovalCard
              language={state.language ?? 'en'}
              pending={state.pendingWrite}
              onApprove={() => controller.approvePendingWrite()}
              onReject={() => controller.rejectPendingWrite()}
            />

            <ShareApprovalCard
              pending={state.pendingShare}
              onApprove={() => controller.approvePendingShare()}
              onReject={() => controller.rejectPendingShare()}
            />
          </section>
        )}

        <QuickActionParamForm
          key={paramFill?.id}
          action={paramFill}
          disabled={actionBlocked}
          onSubmit={onParamFillSubmit}
          onCancel={() => setParamFill(undefined)}
        />

        <Composer
          key={surface}
          busy={state.busy}
          disabled={hasBlockingGate}
          surface={surface}
          allowedIntents={allowedIntents}
          // The plain-string fallback (unreachable while `onInvoke` is set, since Composer prefers it).
          // Still guarded against an unfilled `{{…}}` so this is not a second un-guarded `send` seam if a
          // refactor ever drops `onInvoke` (security review H, LOW).
          onSend={(q) => {
            if (!hasUnfilledPlaceholder(q)) void controller.send(q);
          }}
          onCancel={() => controller.cancel()}
          onInvoke={onInvoke}
          placeholder={SURFACE_PLACEHOLDER[surface]}
          answerOptions={answerOptions}
          mentionOptions={{
            datastore: state.availableDataStores.map((d) => ({
              id: d.resourceName,
              label: d.displayName,
            })),
          }}
        />
      </Toolbar>
    </div>
  );
}

/** The task the pane was closed during: finish only what is left, or dismiss it. */
function ResumeCard({
  request,
  made,
  onFinish,
  onDismiss,
}: {
  request: string;
  made: number;
  onFinish: () => void;
  onDismiss: () => void;
}): JSX.Element {
  const t = uiStrings(detectLanguage(request));
  const shown = request.length > 80 ? `${request.slice(0, 79)}…` : request;
  return (
    <div className="error-card resume-card" role="status">
      <div className="error-card-title">{t.resumeTitle}</div>
      <div className="error-card-body">{t.resumeBody(shown, made)}</div>
      <div className="result-card-actions">
        <button type="button" className="btn sm pr" onClick={onFinish}>
          {t.resumeFinish}
        </button>
        <button type="button" className="btn sm" onClick={onDismiss}>
          {t.resumeDismiss}
        </button>
      </div>
    </div>
  );
}
