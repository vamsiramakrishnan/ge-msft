import { createBridgeDispatch } from '@ge/runtime';
import type {
  ActuationKind,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { buildDocStateSnapshot } from '@ge/content';
import { TEAMS_CAPABILITIES } from './capabilities.js';
import {
  searchTranscript,
  transcriptToContext,
  transcriptToDocStateBlocks,
  type TranscriptInput,
} from './capture.js';
import { planPostMessage } from './actuate-plan.js';
import { meetingEndedEvent, sessionEndEvent, sessionStartEvent } from './events.js';

/**
 * The Teams `DocBridge`. The ONLY place TeamsJS APIs are touched. Like Outlook, Teams is a
 * string-path surface — there is no rich Office.js object model: the bridge captures a transcript
 * window (meeting captions / recent chat turns) + meeting metadata and hands the plain values to
 * `transcriptToContext` (pure, unit-tested). Writes do NOT send silently: a `post-message`
 * actuation stages a reviewable post (or Adaptive Card) through the host's compose/share path,
 * which the user confirms before it leaves. Pure mapping lives in `capture.ts` / `actuate-plan.ts`
 * / `events.ts`; this file is the host wiring.
 *
 * We depend on a tiny feature-detected `TeamsJsLike` interface rather than `@microsoft/teams-js`
 * directly (mirroring web-shell's `MsalLike`), so the core stays unit-testable and dep-light; the
 * app wires a real TeamsJS module (which satisfies this shape) at startup.
 */

export class TeamsBridge implements DocBridge {
  private static readonly dispatcher = createBridgeDispatch<TeamsBridge>(
    'teams',
    {
      'post-message': (host, request) => host.applyPostMessage(request),
    },
    { provenance: 'unsupported' },
  );
  static readonly handledActuations = TeamsBridge.dispatcher.handledActuations;

  readonly surface = 'teams' as const;

  /** Monotonic `<doc_state>` version, bumped on each capture (ADR-0003 Layer B element 1). */
  private docStateVersion = 0;

  constructor(private readonly options: TeamsBridgeOptions = {}) {}

  getCapabilities(): CapabilityManifest {
    return TEAMS_CAPABILITIES;
  }

  /**
   * The one attachable handle is the current transcript window. We surface it only when a
   * transcript snapshot is available (supplied to the constructor or refreshed via
   * `setTranscript`) — Teams gives no native object to enumerate.
   */
  async listContext(): Promise<ContextRef[]> {
    const input = this.transcript;
    if (!input || !input.transcript.trim()) return [];
    const title = input.meetingTitle ? `Transcript: ${input.meetingTitle}` : 'Meeting transcript';
    return [
      {
        id: 'teams:transcript',
        kind: 'transcript',
        surface: 'teams',
        title,
        preview: input.transcript.slice(0, 120),
        live: true,
      },
    ];
  }

  /**
   * Materialize the transcript window into attach-ready context. Re-reads the freshest snapshot
   * (the constructor value, later mutated via `setTranscript`) so a `live` re-resolve at send-time
   * picks up newly captured turns. Empty when no transcript has been captured yet.
   */
  async resolveContext(_ref: ContextRef): Promise<ResolvedContext[]> {
    const input = this.transcript;
    if (!input || !input.transcript.trim()) return [];
    return transcriptToContext(input);
  }

  /**
   * Update the captured transcript window. The host (meeting captions stream / chat read) calls
   * this as new turns arrive; the next `listContext`/`resolveContext` reflects them. Kept off the
   * constructor so the bridge can be created before the transcript is available.
   */
  setTranscript(input: TranscriptInput): void {
    this.options.transcript = input;
  }

  /**
   * ADR-0006 whole-transcript `read` (the runtime's empty-selector read → `captureDocState`): a
   * transcript has no addressable sub-range, so the "document" is the captured window. Builds a
   * snapshot from the meeting title (heading) + bounded turn lines. Reads only the in-memory
   * transcript snapshot (no host round-trip); no transcript captured yet → `undefined`. Version
   * increments per capture.
   */
  async captureDocState(): Promise<DocStateSnapshot | undefined> {
    const input = this.transcript;
    if (!input || !input.transcript.trim()) return undefined;
    const blocks = transcriptToDocStateBlocks(input);
    if (blocks.length === 0) return undefined;
    this.docStateVersion += 1;
    return buildDocStateSnapshot({
      surface: 'teams',
      version: this.docStateVersion,
      ...(input.meetingTitle?.trim() ? { title: input.meetingTitle } : {}),
      blocks,
    });
  }

  /**
   * ADR-0006 `search` read: scan the captured transcript window for `query` and return matching
   * turn lines as `ResolvedContext` data (never instructions), bounded by `searchTranscript`. Scoped
   * to the in-memory transcript window (no cross-meeting read). Empty query / no transcript / no
   * match → `[]`.
   */
  async searchDocument(query: string): Promise<ResolvedContext[]> {
    const q = query.trim();
    const input = this.transcript;
    if (!q || !input || !input.transcript.trim()) return [];
    return searchTranscript(input, q);
  }

  canRevealContext(ref: ContextRef): boolean {
    return (
      ref.surface === 'teams' &&
      ref.kind === 'transcript' &&
      this.resolveDeepLink() !== undefined &&
      typeof this.options.teams?.app?.openLink === 'function'
    );
  }

  async revealContext(ref: ContextRef): Promise<void> {
    if (!this.canRevealContext(ref)) return;
    const link = this.resolveDeepLink();
    const openLink = this.options.teams?.app?.openLink;
    if (!link || typeof openLink !== 'function') return;
    await Promise.resolve(openLink.call(this.options.teams?.app, link));
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    return TeamsBridge.dispatcher.dispatch(this, req);
  }

  /**
   * Stream Teams host lifecycle into the trigger engine, best-effort and feature-detected. We
   * emit `session-start` on registration (the pane is live), and `meeting-ended` when TeamsJS
   * raises its meeting end signal — mapped through the pure builders. `watch` must NEVER throw:
   * a missing TeamsJS module or a failed registration just yields a no-op unsubscribe, and the
   * teardown is wrapped so a host that already removed the handler can't break us.
   */
  watch(emit: (event: HostEvent) => void): Unsubscribe {
    try {
      emit(sessionStartEvent());
    } catch {
      // A consumer throwing on session-start must not abort registration.
    }

    const teams = this.options.teams;
    const onEnd = teams?.meeting?.registerMeetingEndHandler;
    if (!teams || typeof onEnd !== 'function') {
      // No meeting-end source available — session-start already emitted; nothing else to observe.
      return () => {
        this.emitSessionEnd(emit);
      };
    }

    let registered = false;
    try {
      onEnd.call(teams.meeting, () => {
        try {
          const id = this.resolveMeetingId() ?? 'teams:meeting';
          emit(meetingEndedEvent(id));
        } catch {
          // A failed event read must not break the host meeting loop.
        }
      });
      registered = true;
    } catch {
      // Registration failed (older TeamsJS / wrong frame) — degrade to session-start only.
    }

    return () => {
      // TeamsJS exposes no unregister for the meeting-end handler; the unsubscribe is the
      // session-end signal. `registered` is read so the handle reflects what we attached.
      void registered;
      this.emitSessionEnd(emit);
    };
  }

  /** Current best transcript snapshot (constructor value, later updated via `setTranscript`). */
  private get transcript(): TranscriptInput | undefined {
    return this.options.transcript;
  }

  private emitSessionEnd(emit: (event: HostEvent) => void): void {
    try {
      emit(sessionEndEvent());
    } catch {
      // Teardown must stay silent on consumer errors.
    }
  }

  /** Read the opaque meeting/chat id from the host context, defensively narrowing `unknown`. */
  private resolveMeetingId(): string | undefined {
    return this.options.meetingId;
  }

  private resolveDeepLink(): string | undefined {
    const link =
      this.options.deepLink ?? this.options.transcript?.deepLink ?? this.options.meetingId;
    return isSafeTeamsLink(link) ? link : undefined;
  }

  private async applyPostMessage(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planPostMessage(req);
    if (!plan.text.trim() && plan.card === undefined) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_content', message: 'post-message needs params.text or params.html' },
      };
    }

    const compose = resolveComposeTarget(this.options.teams);
    if (!compose) {
      // No host compose path available — degrade to a panel item the user can copy/post manually.
      return {
        ok: true,
        changeId: req.changeId,
        kind: req.kind,
        location: 'panel',
        degraded: true,
      };
    }

    try {
      // Open a reviewable compose/share surface rather than sending — the user confirms before
      // it posts. The staged text/card is what the host renders for review.
      await Promise.resolve(
        compose.fn.call(compose.thisArg, {
          message: plan.text,
          ...(plan.card !== undefined ? { card: plan.card } : {}),
        }),
      );
      return { ok: true, changeId: req.changeId, kind: req.kind, location: 'chat-compose' };
    } catch (err) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'compose_failed', message: errorMessage(err) },
      };
    }
  }
}

/**
 * Constructor options. `transcript`/`meetingId` are pragmatic injection points (Teams has no
 * native model to enumerate), and `teams` is the feature-detected TeamsJS surface — all optional
 * so the bridge can be created before any of them resolve.
 */
export interface TeamsBridgeOptions {
  transcript?: TranscriptInput;
  meetingId?: string;
  /** Explicit deep link to the closest navigable Teams context; not inferred from opaque ids. */
  deepLink?: string;
  teams?: TeamsJsLike;
}

/**
 * The minimal, feature-detected slice of TeamsJS the bridge uses (mirrors web-shell's `MsalLike`).
 * Every member is optional: a real `@microsoft/teams-js` module satisfies this shape, and the
 * bridge narrows each call with `typeof … === 'function'` before invoking it. We only touch:
 *   • `meeting.registerMeetingEndHandler` — the `meeting-ended` event source.
 *   • `chat.openConversation` / `sharing.shareWebContent` — the reviewable compose/share path.
 */
export interface TeamsJsLike {
  app?: {
    openLink?: (deepLink: string) => Promise<unknown> | unknown;
  };
  meeting?: {
    registerMeetingEndHandler?: (handler: () => void) => void;
  };
  chat?: {
    openConversation?: (request: TeamsComposeRequest) => Promise<unknown> | unknown;
  };
  sharing?: {
    shareWebContent?: (request: TeamsComposeRequest) => Promise<unknown> | unknown;
  };
}

/** The staged-for-review payload handed to the host compose/share call. */
export interface TeamsComposeRequest {
  message: string;
  card?: unknown;
}

/** A resolved host compose/share call bound to its owning TeamsJS namespace. */
interface ComposeTarget {
  fn: (request: TeamsComposeRequest) => Promise<unknown> | unknown;
  thisArg: unknown;
}

/**
 * Feature-detect a reviewable compose/share path on the TeamsJS surface: prefer the chat
 * `openConversation` namespace, fall back to `sharing.shareWebContent`. Returns the bound call +
 * its `this` so the host method runs with the right receiver, or undefined when neither exists.
 */
function resolveComposeTarget(teams: TeamsJsLike | undefined): ComposeTarget | undefined {
  const open = teams?.chat?.openConversation;
  if (typeof open === 'function') return { fn: open, thisArg: teams?.chat };
  const share = teams?.sharing?.shareWebContent;
  if (typeof share === 'function') return { fn: share, thisArg: teams?.sharing };
  return undefined;
}

function isSafeTeamsLink(link: string | undefined): link is string {
  return (
    typeof link === 'string' && /^(https:\/\/teams\.microsoft\.com\/|msteams:\/\/)/i.test(link)
  );
}

/** Narrow an unknown thrown value to a message string without assuming `Error`. */
function errorMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (typeof err === 'string') return err;
  return 'Teams compose failed';
}

/** Actual dispatch keys; conformance checks these against the advertised capabilities. */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = TeamsBridge.handledActuations;
