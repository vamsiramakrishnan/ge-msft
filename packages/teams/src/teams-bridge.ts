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
import type { AttachmentBytes, DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import {
  attachmentToContext,
  buildDocStateSnapshot,
  classifyAttachment,
  MAX_INLINE_ATTACHMENT_BYTES,
} from '@ge/content';
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
    const refs: ContextRef[] = [];
    const input = this.transcript;
    if (input && input.transcript.trim()) {
      const title = input.meetingTitle ? `Transcript: ${input.meetingTitle}` : 'Meeting transcript';
      refs.push({
        id: 'teams:transcript',
        kind: 'transcript',
        surface: 'teams',
        title,
        preview: input.transcript.slice(0, 120),
        live: true,
      });
    }
    refs.push(...(await this.listSharedFileRefs()));
    return refs;
  }

  /**
   * Materialize a handle into attach-ready context. The transcript re-reads the freshest snapshot
   * (the constructor value, later mutated via `setTranscript`) so a `live` re-resolve at send-time
   * picks up newly captured turns. A shared file goes through the shared attachment pipeline:
   * text-like files are read (bounded) and inlined as data; PDF/Office files become one uploadable
   * `file` part the runtime materializes on an explicit attach. Empty when nothing is captured.
   */
  async resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    if (ref.kind === 'attachment') return this.resolveSharedFile(ref);
    const input = this.transcript;
    if (!input || !input.transcript.trim()) return [];
    return transcriptToContext(input);
  }

  /**
   * Raw bytes of a shared file for the runtime's session-context upload (explicit user attach
   * only). Reads as the signed-in user through the injected source; never retained here.
   */
  async readAttachment(ref: ContextRef): Promise<AttachmentBytes> {
    const file = this.sharedFileFor(ref);
    const source = this.options.files;
    if (!file || !source) throw new Error('This file is not available in the current chat.');
    const read = await source.read(file.id);
    const { mimeType } = classifyAttachment(file.name, read.mimeType ?? file.mimeType);
    return { fileName: file.name, mimeType, bytes: read.bytes };
  }

  /**
   * Install (or replace) the source of files shared in the current chat/meeting. The host wires
   * this after sign-in, when a delegated Graph reader for the chat exists; without it the bridge
   * lists no files. Kept off the constructor like `setTranscript`.
   */
  setSharedFiles(source: SharedFilesSource | undefined): void {
    this.options.files = source;
    this.sharedFiles = undefined;
  }

  private sharedFiles: SharedFile[] | undefined;

  private async listSharedFileRefs(): Promise<ContextRef[]> {
    const source = this.options.files;
    if (!source) return [];
    let files: SharedFile[];
    try {
      files = await source.list();
    } catch {
      // A chat the user cannot read, or no chat in this frame: no files, never a broken pane.
      return [];
    }
    this.sharedFiles = files;
    return files.map((file): ContextRef => {
      const { cls, mimeType } = classifyAttachment(file.name, file.mimeType);
      const how =
        cls === 'text'
          ? 'readable text'
          : cls === 'file'
            ? 'attach to read as a context file'
            : 'content not readable';
      const who = file.sharedBy ? ` · shared by ${file.sharedBy}` : '';
      return {
        id: sharedFileRefId(file.id),
        kind: 'attachment',
        surface: 'teams',
        title: file.name,
        preview: `${how}${who}`,
        mimeType,
        ...(file.sizeBytes !== undefined ? { sizeBytes: file.sizeBytes } : {}),
        ...(file.webUrl ? { hostRef: { type: 'teams.deepLink', url: file.webUrl } } : {}),
      };
    });
  }

  private sharedFileFor(ref: ContextRef): SharedFile | undefined {
    const id = sharedFileIdFromRef(ref);
    return id === undefined ? undefined : this.sharedFiles?.find((f) => f.id === id);
  }

  private async resolveSharedFile(ref: ContextRef): Promise<ResolvedContext[]> {
    const source = this.options.files;
    if (!source) return [];
    if (!this.sharedFiles) await this.listSharedFileRefs();
    const file = this.sharedFileFor(ref);
    if (!file) return [];
    const base = {
      sourceId: ref.id,
      surface: 'teams' as const,
      fileName: file.name,
      ...(file.mimeType ? { mimeType: file.mimeType } : {}),
      ...(file.sizeBytes !== undefined ? { sizeBytes: file.sizeBytes } : {}),
    };
    // Only text-like files within the inline bound are read here; a binary is uploaded by the
    // runtime on attach instead, and an oversized text file grounds by name.
    if (
      classifyAttachment(file.name, file.mimeType).cls !== 'text' ||
      (file.sizeBytes ?? 0) > MAX_INLINE_ATTACHMENT_BYTES
    )
      return attachmentToContext(base);
    try {
      const read = await source.read(file.id);
      return attachmentToContext({ ...base, bytes: read.bytes });
    } catch {
      return attachmentToContext(base);
    }
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
  /** Files shared in the current chat/meeting, read as the signed-in user (wired by the host). */
  files?: SharedFilesSource;
}

/** One file shared in the current chat or meeting, as the host's file source lists it. */
export interface SharedFile {
  id: string;
  name: string;
  mimeType?: string;
  sizeBytes?: number;
  webUrl?: string;
  sharedBy?: string;
}

/**
 * The host-injected reader for files shared in the current chat. The bridge never calls Graph
 * itself (Plane B stays in `@ge/graph-client`, wired from the shell); it only asks this source,
 * which must act as the signed-in user and may refuse any file the user cannot open.
 */
export interface SharedFilesSource {
  list(): Promise<SharedFile[]>;
  read(id: string): Promise<{ bytes: Uint8Array; mimeType?: string }>;
}

const SHARED_FILE_PREFIX = 'teams:file:';

export function sharedFileRefId(fileId: string): string {
  return `${SHARED_FILE_PREFIX}${fileId}`;
}

function sharedFileIdFromRef(ref: ContextRef): string | undefined {
  return ref.kind === 'attachment' && ref.id.startsWith(SHARED_FILE_PREFIX)
    ? ref.id.slice(SHARED_FILE_PREFIX.length)
    : undefined;
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
