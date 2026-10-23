/**
 * Live wiring for scenario runs: the real Gemini Enterprise engine from `packages/web-shell/.env`,
 * signed in with the developer's `gcloud` account instead of the browser's Entra → WIF exchange.
 * Everything the pane would send still goes through the real client; only the token source and
 * the transcript recorder are swapped in. Test environments only.
 */
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { AuthClient } from '@ge/runtime';
import type { Brand } from '../brand.js';

export interface TranscriptEntry {
  at: string;
  url: string;
  status: number;
  request: string;
  /** The model's visible text for this call, joined from the streamed replies. */
  reply: string;
  /** Distinguish an active or failed stream from a completed response with an empty answer. */
  bodyState?: 'streaming' | 'complete' | 'failed';
}

const REPO_ROOT = join(__dirname, '..', '..', '..', '..');

/**
 * The build's brand (GE_BRAND, else `brands/current`), read from this repo the way the web build
 * applies it, so a scenario sends the same name line and writing rules users of that build get.
 */
export function scenarioBrand(): Brand {
  const pinned = join(REPO_ROOT, 'brands', 'current');
  const id =
    process.env.GE_BRAND?.trim() ||
    (existsSync(pinned) ? readFileSync(pinned, 'utf8').trim() : '') ||
    'default';
  if (!/^[a-z0-9][a-z0-9-]*$/.test(id)) throw new Error(`GE_BRAND "${id}" is not a brand id.`);
  const raw = JSON.parse(
    readFileSync(join(REPO_ROOT, 'brands', id, 'brand.json'), 'utf8'),
  ) as Brand;
  // The fields the pane reads from its build-time brand (see tools/brand/brand.mjs `brandText`).
  return {
    name: raw.name,
    assistantName: raw.assistantName,
    fullName: raw.fullName,
    ...(raw.tagline ? { tagline: raw.tagline } : {}),
    ...(raw.promptName ? { promptName: raw.promptName } : {}),
    ...(raw.promptRules?.length ? { promptRules: raw.promptRules } : {}),
    ...accent(id),
  };
}

/** The brand's main colour, as the build reads it (tools/brand/brand.mjs `primaryColor`). */
function accent(id: string): { accentColor?: string } {
  const theme = join(REPO_ROOT, 'brands', id, 'theme.css');
  if (!existsSync(theme)) return {};
  const css = readFileSync(theme, 'utf8');
  const value = (name: string): string | undefined =>
    new RegExp(`--${name}\\s*:\\s*([^;]+);`).exec(css)?.[1]?.trim();
  let v = value('primary');
  const ref = v ? /^var\(--([\w-]+)\)$/.exec(v) : null;
  if (ref) v = value(ref[1]!);
  return v && /^#[0-9a-fA-F]{6}$/.test(v) ? { accentColor: v.toUpperCase() } : {};
}

/** `packages/web-shell/.env`, then the process environment (same order as the Vite build). */
export function loadShellEnv(): Record<string, string> {
  const env: Record<string, string> = {};
  const file = join(REPO_ROOT, 'packages', 'web-shell', '.env');
  if (existsSync(file)) {
    for (const raw of readFileSync(file, 'utf8').split(/\r?\n/)) {
      const m = /^(?:export\s+)?([A-Z0-9_]+)\s*=\s*(.*)$/.exec(raw.trim());
      if (m) env[m[1]!] = m[2]!.replace(/^['"]|['"]$/g, '');
    }
  }
  for (const [key, value] of Object.entries(process.env)) {
    if (key.startsWith('VITE_') && value !== undefined) env[key] = value;
  }
  return env;
}

let cachedToken: { value: string; at: number } | undefined;

/**
 * A Google access token from `gcloud` (or GE_ACCESS_TOKEN), refreshed every 5 minutes. `gcloud`
 * hands back its own cached token, which can already be most of an hour old: with a 45-minute
 * refresh, a long run got HTTP 401 on every request after it expired (2026-10-19, 14 scenarios).
 */
export function googleAccessToken(): string {
  const inline = process.env.GE_ACCESS_TOKEN;
  if (inline) return inline;
  if (!cachedToken || Date.now() - cachedToken.at > 5 * 60_000) {
    const value = execFileSync('gcloud', ['auth', 'print-access-token'], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    }).trim();
    cachedToken = { value, at: Date.now() };
  }
  return cachedToken.value;
}

/** Stand-in for the signed-in Microsoft user; the Google token comes from {@link liveFetch}. */
export const scenarioAuth: AuthClient = {
  getIdToken: async () => 'scenario-runner-id-token',
  getIdentity: async () => ({ username: 'scenario-runner@test', displayName: 'Scenario runner' }),
};

/** Visible model text from a streamAssist JSON-array body (best effort, for the transcript). */
function replyText(body: string): string {
  try {
    const chunks = JSON.parse(body) as Array<{
      answer?: { replies?: Array<{ groundedContent?: { content?: { text?: string } } }> };
    }>;
    return chunks
      .flatMap((c) => c.answer?.replies ?? [])
      .map((r) => r.groundedContent?.content?.text ?? '')
      .join('');
  } catch {
    return body.slice(0, 2000);
  }
}

/**
 * The fetch the scenario session uses: answers the STS exchange with the gcloud token, adds the
 * quota project to Google API calls, and records each Gemini request/reply in `transcript`.
 */
export function liveFetch(
  env: Record<string, string>,
  transcript: TranscriptEntry[],
): typeof fetch {
  const quotaProject = env.VITE_WIF_USER_PROJECT || env.VITE_GCP_PROJECT;
  return (async (input: RequestInfo | URL, init: RequestInit = {}) => {
    const url = String(input instanceof Request ? input.url : input);
    if (url.startsWith('https://sts.googleapis.com/')) {
      return new Response(
        JSON.stringify({
          access_token: googleAccessToken(),
          token_type: 'Bearer',
          issued_token_type: 'urn:ietf:params:oauth:token-type:access_token',
          expires_in: 2700,
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      );
    }
    const headers = new Headers(init.headers);
    if (quotaProject && url.includes('googleapis.com')) {
      headers.set('X-Goog-User-Project', quotaProject);
    }
    const res = await fetch(input, { ...init, headers });
    if (url.includes(':streamAssist')) {
      const entry: TranscriptEntry = {
        at: new Date().toISOString(),
        url: url.replace(/projects\/[^/]+/, 'projects/…'),
        status: res.status,
        request: typeof init.body === 'string' ? init.body : '',
        reply: '',
        bodyState: 'streaming',
      };
      transcript.push(entry);
      // Return headers immediately, as real fetch does. Awaiting the whole clone here makes
      // an active stream look like a stalled POST and hides thought/heartbeat frames from the client.
      void res
        .clone()
        .text()
        .then(
          (body) => {
            entry.reply = replyText(body);
            entry.bodyState = 'complete';
          },
          () => {
            entry.bodyState = 'failed';
          },
        );
    }
    return res;
  }) as typeof fetch;
}

/**
 * Replays a recorded scenario run: each streamAssist call gets the next recorded reply, in order,
 * and nothing reaches Google. Reproduces a live bug (and checks its fix) at no Gemini cost. If the
 * code asks the model more often than the recording did, the extra turns get an empty reply.
 */
export function replayFetch(
  recordedReplies: string[],
  transcript: TranscriptEntry[],
): typeof fetch {
  const replies = [...recordedReplies];
  return (async (input: RequestInfo | URL, init: RequestInit = {}) => {
    const url = String(input instanceof Request ? input.url : input);
    if (url.includes(':streamAssist')) {
      const text = replies.shift() ?? '';
      transcript.push({
        at: new Date().toISOString(),
        url: url.replace(/projects\/[^/]+/, 'projects/…'),
        status: 200,
        request: typeof init.body === 'string' ? init.body : '',
        reply: text,
      });
      // The transcript joins thought summaries ("**Analyzing…**") into the reply text; live, they
      // arrive as separate thought parts. Send leading summary lines back as thoughts, or a reply
      // like `**Heading**\n```cmd\ndone "…"` replays as prose with no command block.
      const thoughts = /^(?:[^\S\n]*\*\*[^\n]*\*\*[^\S\n]*\n)+/.exec(text)?.[0] ?? '';
      const parts = [
        ...(thoughts ? [{ groundedContent: { content: { text: thoughts, thought: true } } }] : []),
        { groundedContent: { content: { text: text.slice(thoughts.length) } } },
      ];
      const chunk = {
        answer: { state: 'SUCCEEDED', replies: parts },
        sessionInfo: { session: 'projects/replay/locations/global/sessions/replay' },
      };
      return new Response(JSON.stringify([chunk]), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      });
    }
    if (url.startsWith('https://sts.googleapis.com/'))
      return new Response(
        JSON.stringify({
          access_token: 'replay-token',
          token_type: 'Bearer',
          issued_token_type: 'urn:ietf:params:oauth:token-type:access_token',
          expires_in: 2700,
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      );
    return new Response('{}', { status: 200, headers: { 'content-type': 'application/json' } });
  }) as typeof fetch;
}
