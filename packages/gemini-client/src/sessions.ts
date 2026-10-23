import { z } from 'zod';
import { defaultFetch, getJson, type FetchLike } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';
import { sessionUrl, sessionsUrl, type GeminiClientConfig } from './config.js';

export interface ConversationSummary {
  name: string;
  id: string;
  title: string;
  turnCount: number;
  isPinned: boolean;
  state?: string;
  startedAt?: string;
  endedAt?: string;
  updatedAt?: string;
  /** The Office app the chat ran in, read from the add-in's `<doc_state surface=…>` context part. */
  surface?: string;
}

export interface ConversationListResult {
  conversations: ConversationSummary[];
  nextPageToken?: string;
}

export interface ConversationSession extends ConversationSummary {
  turns: Array<{
    queryText?: string;
    /** The user's own words: the query parts without a MIME type (context parts carry one). */
    userText?: string;
    createTime?: string;
    answerState?: string;
    /** The visible answer: non-thought reply text, joined (only with `includeAnswerDetails`). */
    answerText?: string;
    /** The answer also had a file or code reply (e.g. a chart) that text cannot show. */
    answerHasMedia?: boolean;
  }>;
}

const QueryPartSchema = z
  .object({
    text: z.string().optional(),
    mimeType: z.string().optional(),
  })
  .passthrough();

const QuerySchema = z
  .object({
    text: z.string().optional(),
    createTime: z.string().optional(),
    parts: z.array(QueryPartSchema).optional(),
  })
  .passthrough();

const AssistReplySchema = z
  .object({
    groundedContent: z
      .object({
        content: z
          .object({ text: z.string().optional(), thought: z.boolean().optional() })
          .passthrough()
          .optional(),
      })
      .passthrough()
      .optional(),
  })
  .passthrough();

const SessionTurnSchema = z
  .object({
    query: QuerySchema.optional(),
    createdAt: z.string().optional(),
    detailedAssistAnswer: z
      .object({ state: z.string().optional(), replies: z.array(AssistReplySchema).optional() })
      .passthrough()
      .optional(),
    detailedAnswer: z.object({ state: z.string().optional() }).passthrough().optional(),
  })
  .passthrough();

const SessionSchema = z
  .object({
    name: z.string(),
    displayName: z.string().optional(),
    startTime: z.string().optional(),
    endTime: z.string().optional(),
    updateTime: z.string().optional(),
    isPinned: z.boolean().optional(),
    state: z.string().optional(),
    turns: z.array(SessionTurnSchema).optional(),
  })
  .passthrough();

const ListSessionsResponseSchema = z
  .object({
    sessions: z.array(SessionSchema).optional(),
    nextPageToken: z.string().optional(),
  })
  .passthrough();

export class ConversationClient {
  constructor(
    private readonly tokens: TokenSource,
    private readonly config: GeminiClientConfig,
    private readonly fetchImpl: FetchLike = defaultFetch,
  ) {}

  async listConversations(
    opts: { pageSize?: number; pageToken?: string; signal?: AbortSignal } = {},
  ): Promise<ConversationListResult> {
    const url = new URL(sessionsUrl(this.config));
    url.searchParams.set('pageSize', String(clampPageSize(opts.pageSize)));
    if (opts.pageToken) url.searchParams.set('pageToken', opts.pageToken);
    const raw = await getJson(url.toString(), this.tokens, this.fetchImpl, opts.signal);
    const parsed = ListSessionsResponseSchema.parse(raw);
    const conversations = (parsed.sessions ?? []).map(toSummary);
    return {
      conversations,
      ...(parsed.nextPageToken ? { nextPageToken: parsed.nextPageToken } : {}),
    };
  }

  async getConversation(
    sessionIdOrName: string,
    opts: { includeAnswerDetails?: boolean; signal?: AbortSignal } = {},
  ): Promise<ConversationSession> {
    const url = new URL(sessionUrl(this.config, sessionIdOrName));
    if (opts.includeAnswerDetails) url.searchParams.set('includeAnswerDetails', 'true');
    const raw = await getJson(url.toString(), this.tokens, this.fetchImpl, opts.signal);
    const parsed = SessionSchema.parse(raw);
    return {
      ...toSummary(parsed),
      turns: (parsed.turns ?? []).map((turn) => ({
        ...(queryText(turn.query) ? { queryText: queryText(turn.query) } : {}),
        ...(userText(turn.query) ? { userText: userText(turn.query) } : {}),
        ...((turn.query?.createTime ?? turn.createdAt)
          ? { createTime: turn.query?.createTime ?? turn.createdAt }
          : {}),
        ...answerOf(turn),
        ...((turn.detailedAssistAnswer?.state ?? turn.detailedAnswer?.state)
          ? { answerState: turn.detailedAssistAnswer?.state ?? turn.detailedAnswer?.state }
          : {}),
      })),
    };
  }
}

function toSummary(session: z.infer<typeof SessionSchema>): ConversationSummary {
  const turns = session.turns ?? [];
  const lastQuery = [...turns].reverse().find((turn) => queryText(turn.query))?.query;
  const lastTurnAt = [...turns].reverse().find((turn) => turn.createdAt)?.createdAt;
  const updatedAt =
    session.updateTime ??
    lastQuery?.createTime ??
    lastTurnAt ??
    session.endTime ??
    session.startTime;
  const id = session.name.split('/').pop() ?? session.name;
  const surface = surfaceOf(turns);
  return {
    name: session.name,
    id,
    title: session.displayName?.trim() || userText(turns[0]?.query) || id,
    turnCount: turns.length,
    isPinned: session.isPinned ?? false,
    ...(session.state ? { state: session.state } : {}),
    ...(session.startTime ? { startedAt: session.startTime } : {}),
    ...(session.endTime ? { endedAt: session.endTime } : {}),
    ...(updatedAt ? { updatedAt } : {}),
    ...(surface ? { surface } : {}),
  };
}

/** The first `<doc_state surface=…>` part the add-in sent; chats from elsewhere have none. */
function surfaceOf(turns: Array<z.infer<typeof SessionTurnSchema>>): string | undefined {
  for (const turn of turns) {
    for (const part of turn.query?.parts ?? []) {
      const match = /^<doc_state surface=([a-z]+)\b/.exec(part.text ?? '');
      if (match) return match[1];
    }
  }
  return undefined;
}

function queryText(query: z.infer<typeof QuerySchema> | undefined): string | undefined {
  if (!query) return undefined;
  const direct = query.text?.trim();
  if (direct) return direct;
  const parts = (query.parts ?? [])
    .map((part) => part.text?.trim())
    .filter((text): text is string => Boolean(text));
  return parts.length ? parts.join(' ') : undefined;
}

/** Context the add-in sends, in case a part ever arrives without its MIME type. */
const CONTEXT_PART = /^(?:<doc_state\b|Working-document read\b|Working context so far\b)/;

/** Marks the pane's own instructions (brand name, writing rules) inside a query. */
export const PANE_INSTRUCTIONS_OPEN = '<pane_instructions>';
export const PANE_INSTRUCTIONS_CLOSE = '</pane_instructions>';
const PANE_BLOCK = /^\s*<pane_instructions>[\s\S]*?<\/pane_instructions>\s*/;
/** The unmarked name line sent before 2026-10-16 ("Your name is ContosoAI. …"), then a blank line. */
const LEGACY_NAME_LINE = /^\s*Your name is [^\n]*(?:\n[^\n]+)*\n\n/;
/**
 * The same line sent as its own part (with attached context, e.g. every Outlook turn): no blank
 * line follows it (UAT 2026-10-16: shown above the user's question).
 */
const LEGACY_NAME_PART =
  /^\s*Your name is [^\n]+? only; never call yourself Gemini or Gemini Enterprise\.\s*$/;

/** A query text without the pane's instruction block (or the older unmarked name line). */
export function withoutPaneInstructions(text: string): string {
  return text.replace(PANE_BLOCK, '').replace(LEGACY_NAME_LINE, '').replace(LEGACY_NAME_PART, '');
}

/**
 * The user's own words. The add-in sends context (the `<doc_state>` snapshot, attached reads) as
 * parts with a MIME type and the typed question as a bare text part.
 */
function userText(query: z.infer<typeof QuerySchema> | undefined): string | undefined {
  if (!query) return undefined;
  if (!query.parts?.length) return withoutPaneInstructions(query.text ?? '').trim() || undefined;
  const own = query.parts
    .filter((part) => !part.mimeType)
    .map((part) => withoutPaneInstructions(part.text ?? '').trim())
    .filter((text): text is string => Boolean(text) && !CONTEXT_PART.test(text!));
  return own.length ? own.join('\n') : undefined;
}

/** The answer as the user saw it: reply text without the model's thoughts. */
function answerOf(turn: z.infer<typeof SessionTurnSchema>): {
  answerText?: string;
  answerHasMedia?: boolean;
} {
  const replies = turn.detailedAssistAnswer?.replies ?? [];
  const content = replies.map((reply) => reply.groundedContent?.content);
  const text = content
    .filter((c) => c?.text && c.thought !== true)
    .map((c) => c!.text!)
    .join('')
    .trim();
  const hasMedia = content.some(
    (c) => c && c.thought !== true && !c.text && ('file' in c || 'executableCode' in c),
  );
  return { ...(text ? { answerText: text } : {}), ...(hasMedia ? { answerHasMedia: true } : {}) };
}

function clampPageSize(pageSize: number | undefined): number {
  if (!Number.isFinite(pageSize)) return 20;
  return Math.min(50, Math.max(1, Math.trunc(pageSize!)));
}
