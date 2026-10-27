import type { TeamsBridgeOptions, TeamsJsLike } from '@ge/teams';
import type { TeamsContextLike } from './teams-shared-files.js';

/**
 * TeamsJS for the Teams task pane. The SDK is loaded only on the Teams surface (a dynamic import,
 * so Office hosts never ship it) and initialized with a bound: `app.initialize()` never resolves
 * outside a Teams frame, and a pane that cannot reach the host must still boot. What the shell
 * hands on is deliberately narrow — the context reader for the chat id, `app.openLink` and the
 * meeting-end handler — which is exactly the `TeamsJsLike` slice the bridge feature-detects today.
 */
export interface TeamsSdk {
  app: {
    initialize(): Promise<void>;
    getContext(): Promise<TeamsContextLike>;
    openLink?: (deepLink: string) => Promise<unknown> | unknown;
  };
  meeting?: { registerMeetingEndHandler?: (handler: () => void) => void };
}

export const TEAMS_INIT_TIMEOUT_MS = 5_000;

export async function loadTeamsSdk(
  importer: () => Promise<TeamsSdk> = () => import('@microsoft/teams-js') as Promise<TeamsSdk>,
  timeoutMs = TEAMS_INIT_TIMEOUT_MS,
): Promise<TeamsSdk | undefined> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    const sdk = await importer();
    await Promise.race([
      sdk.app.initialize(),
      new Promise<never>((_, reject) => {
        timer = setTimeout(
          () => reject(new Error('TeamsJS did not initialize in time.')),
          timeoutMs,
        );
      }),
    ]);
    return sdk;
  } catch {
    return undefined;
  } finally {
    clearTimeout(timer);
  }
}

/** The bridge's feature-detected slice, built from an initialized SDK. */
export function teamsBridgeOptions(sdk: TeamsSdk | undefined): TeamsBridgeOptions | undefined {
  if (!sdk) return undefined;
  const teams: TeamsJsLike = {
    ...(sdk.app.openLink ? { app: { openLink: sdk.app.openLink.bind(sdk.app) } } : {}),
    ...(sdk.meeting?.registerMeetingEndHandler
      ? {
          meeting: {
            registerMeetingEndHandler: sdk.meeting.registerMeetingEndHandler.bind(sdk.meeting),
          },
        }
      : {}),
  };
  return { teams };
}
