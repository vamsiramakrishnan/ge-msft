/**
 * In-memory **Teams host simulator**. Unlike Word/Excel/Outlook there is no rich Office.js object
 * model and no `globalThis` namespace to install: TeamsJS is **feature-detected and injected** into
 * the {@link "@ge/teams"!TeamsBridge} constructor (mirroring web-shell's `MsalLike`). So the "host"
 * the bridge drives is exactly the small `TeamsJsLike` slice plus the transcript snapshot the host
 * captures from the meeting captions / chat read — and this simulator fakes only that slice.
 *
 * Enumerated host calls modelled (the fidelity boundary for Teams — see `TeamsJsLike`):
 *   - `meeting.registerMeetingEndHandler(handler)` — the `meeting-ended` event source. The handler
 *     the bridge registers is captured so a test can `fireMeetingEnd()` to drive `watch()`.
 *   - `chat.openConversation({ message, card? })` — the reviewable compose path. The staged post is
 *     RECORDED (never "sent"): it mirrors the host opening a compose/share surface the user confirms.
 *   - `sharing.shareWebContent({ message, card? })` — the fallback reviewable share path, recorded
 *     the same way. A seed can choose which path the host exposes.
 *
 * Fidelity notes / boundary:
 *   - The compose call RECORDS the staged post and resolves — exactly the host's reviewable-compose
 *     contract: nothing leaves the client without the user. A test asserts the post was *staged*
 *     (composed for review), never auto-sent, by reading the recorded posts back.
 *   - The transcript window is UNTRUSTED host content. It is handed to the bridge as a plain
 *     `TranscriptInput` string + metadata and flows through `@ge/content` as DATA, never as
 *     instructions; a test can seed a transcript carrying an injection attempt and assert it is
 *     surfaced verbatim as context, not executed.
 */

import type { TeamsBridgeOptions, TeamsComposeRequest, TeamsJsLike } from '@ge/teams';

/** A recorded post the host staged for review via `chat.openConversation` / `sharing.shareWebContent`. */
export interface StagedPost {
  /** Which feature-detected path the bridge took. */
  via: 'chat' | 'sharing';
  /** The reviewable message body the host rendered for the user to confirm. */
  message: string;
  /** The optional Adaptive Card payload, when the post carried one. */
  card?: unknown;
}

/** Which reviewable compose path the fake host exposes (mirrors a real client's capabilities). */
export type ComposePath = 'chat' | 'sharing';

/** The Teams host seed: the transcript window, the meeting id, and which compose path is available. */
export interface TeamsSeed {
  meetingTitle?: string;
  transcript: string;
  participants?: string[];
  /** The opaque meeting id the host carries into a `meeting-ended` event. */
  meetingId?: string;
  /** Which reviewable compose surface the host exposes. Defaults to `chat`. */
  composePath?: ComposePath;
  /** When false, the host exposes NO meeting-end source (older TeamsJS / wrong frame). */
  meetingEndSource?: boolean;
  /** When set, the chosen compose path throws this on invocation (host compose failure path). */
  composeThrows?: Error;
}

/** A read-back view of the Teams host after a run. */
export interface TeamsSnapshot {
  /** Every post the host staged for review, in order. */
  stagedPosts: ReadonlyArray<StagedPost>;
  /** Whether the bridge registered a meeting-end handler. */
  meetingEndRegistered: boolean;
}

/** The installed Teams simulator. */
export interface TeamsSimulator {
  readonly seed: TeamsSeed;
  /**
   * The feature-detected TeamsJS surface + transcript snapshot to hand the bridge constructor.
   * (Teams injects rather than installing a `globalThis` namespace.)
   */
  readonly bridgeOptions: TeamsBridgeOptions;
  /** Fire the TeamsJS meeting-end signal (drives the bridge's `watch()` end handler). */
  fireMeetingEnd(): void;
  /** Whether the bridge registered a meeting-end handler yet. */
  meetingEndRegistered(): boolean;
  snapshot(): TeamsSnapshot;
  /** Teardown — Teams installs no globals, so this is a no-op kept for harness symmetry. */
  restore(): void;
}

/**
 * Build a {@link TeamsSeed} from a transcript window, defaulting the compose path to `chat` and the
 * meeting-end source to present.
 */
export function teamsSeed(init: {
  meetingTitle?: string;
  transcript: string;
  participants?: string[];
  meetingId?: string;
  composePath?: ComposePath;
  meetingEndSource?: boolean;
  composeThrows?: Error;
}): TeamsSeed {
  return {
    ...(init.meetingTitle !== undefined ? { meetingTitle: init.meetingTitle } : {}),
    transcript: init.transcript,
    ...(init.participants ? { participants: init.participants } : {}),
    ...(init.meetingId !== undefined ? { meetingId: init.meetingId } : {}),
    composePath: init.composePath ?? 'chat',
    meetingEndSource: init.meetingEndSource ?? true,
    ...(init.composeThrows ? { composeThrows: init.composeThrows } : {}),
  };
}

/** A realistic meeting fixture: a renewal sync whose transcript carries action items to post back. */
export function defaultTeamsSeed(): TeamsSeed {
  return teamsSeed({
    meetingTitle: 'Q3 Renewal Sync',
    meetingId: '19:meeting_abc@thread.v2',
    participants: ['Pat (AE)', 'Sam (Procurement)'],
    transcript: [
      'Pat: We can hold the SLA at 99.5% for the renewal.',
      'Sam: The liability cap needs to stay at 6 months of fees.',
      'Pat: Agreed. I will send the addendum by Friday.',
    ].join('\n'),
  });
}

/**
 * Install an in-memory Teams host: builds the feature-detected `TeamsJsLike` surface + transcript
 * snapshot the REAL {@link "@ge/teams"!TeamsBridge} constructor consumes, recording every staged
 * (reviewable) post and capturing the meeting-end handler the bridge registers. No `globalThis`
 * namespace is installed — Teams injects its host slice — so this needs no `restore` of globals.
 */
export function installFakeTeams(seed: TeamsSeed = defaultTeamsSeed()): TeamsSimulator {
  const stagedPosts: StagedPost[] = [];
  let endHandler: (() => void) | undefined;
  let registered = false;

  const stage =
    (via: ComposePath) =>
    (req: TeamsComposeRequest): Promise<unknown> => {
      if (seed.composeThrows) throw seed.composeThrows;
      stagedPosts.push({
        via,
        message: req.message,
        ...(req.card !== undefined ? { card: req.card } : {}),
      });
      return Promise.resolve(undefined);
    };

  const teams: TeamsJsLike = {
    ...(seed.meetingEndSource !== false
      ? {
          meeting: {
            registerMeetingEndHandler: (handler: () => void): void => {
              endHandler = handler;
              registered = true;
            },
          },
        }
      : {}),
    ...(seed.composePath === 'sharing'
      ? { sharing: { shareWebContent: stage('sharing') } }
      : { chat: { openConversation: stage('chat') } }),
  };

  const bridgeOptions: TeamsBridgeOptions = {
    teams,
    transcript: {
      ...(seed.meetingTitle !== undefined ? { meetingTitle: seed.meetingTitle } : {}),
      transcript: seed.transcript,
      ...(seed.participants ? { participants: seed.participants } : {}),
    },
    ...(seed.meetingId !== undefined ? { meetingId: seed.meetingId } : {}),
  };

  return {
    seed,
    bridgeOptions,
    fireMeetingEnd: () => endHandler?.(),
    meetingEndRegistered: () => registered,
    snapshot: () => ({
      stagedPosts: stagedPosts.map((p) => ({ ...p })),
      meetingEndRegistered: registered,
    }),
    restore: () => {
      /* Teams installs no globals; nothing to restore. */
    },
  };
}
