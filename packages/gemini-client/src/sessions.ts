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
}

export interface ConversationListResult {
  conversations: ConversationSummary[];
  nextPageToken?: string;
}

export interface ConversationSession extends ConversationSummary {
  turns: Array<{
    queryText?: string;
    createTime?: string;
    answerState?: string;
  }>;
}

const QueryPartSchema = z
  .object({
    text: z.string().optional(),
  })
  .passthrough();

const QuerySchema = z
  .object({
    text: z.string().optional(),
    createTime: z.string().optional(),
    parts: z.array(QueryPartSchema).optional(),
  })
  .passthrough();

const SessionTurnSchema = z
  .object({
    query: QuerySchema.optional(),
    detailedAssistAnswer: z.object({ state: z.string().optional() }).passthrough().optional(),
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
        ...(turn.query?.createTime ? { createTime: turn.query.createTime } : {}),
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
  const updatedAt =
    session.updateTime ?? lastQuery?.createTime ?? session.endTime ?? session.startTime;
  const id = session.name.split('/').pop() ?? session.name;
  return {
    name: session.name,
    id,
    title: session.displayName?.trim() || queryText(turns[0]?.query) || id,
    turnCount: turns.length,
    isPinned: session.isPinned ?? false,
    ...(session.state ? { state: session.state } : {}),
    ...(session.startTime ? { startedAt: session.startTime } : {}),
    ...(session.endTime ? { endedAt: session.endTime } : {}),
    ...(updatedAt ? { updatedAt } : {}),
  };
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

function clampPageSize(pageSize: number | undefined): number {
  if (!Number.isFinite(pageSize)) return 20;
  return Math.min(50, Math.max(1, Math.trunc(pageSize!)));
}
