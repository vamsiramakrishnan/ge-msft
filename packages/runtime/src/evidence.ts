import type { ResolvedContext } from '@ge/contracts';
import type {
  SearchClient,
  RankClient,
  GroundingClient,
  SearchHit,
  GroundingFact,
} from '@ge/gemini-client';
import type { RuntimeHooks } from './hooks.js';

export interface EvidenceState {
  status: 'idle' | 'searching' | 'ready' | 'checking' | 'checked' | 'unavailable';
  sources: Array<{ id: string; title: string; uri?: string }>;
  score?: number;
  checkedClaims?: number;
  message: string;
}
export interface EvidenceOptions {
  search: Pick<SearchClient, 'search'>;
  rank: Pick<RankClient, 'rank'>;
  grounding: Pick<GroundingClient, 'check'>;
  /** Tenant-controlled gate. Scores express support in retrieved snippets, not proof of correctness. */
  requiredSupport?: number;
}
const safeUri = (uri: string | undefined): string | undefined => {
  try {
    const url = new URL(uri ?? '');
    return ['https:', 'http:'].includes(url.protocol) ? url.href : undefined;
  } catch {
    return undefined;
  }
};

/** Scope first, retrieve, deduplicate, rerank, frame as untrusted evidence, then check claims. */
export class EvidencePipeline {
  private current: EvidenceState = {
    status: 'idle',
    sources: [],
    message: 'Select a data store to assemble evidence.',
  };
  private readonly tasks = new Map<string, GroundingFact[]>();
  constructor(private readonly options: EvidenceOptions) {
    if (
      options.requiredSupport !== undefined &&
      (!Number.isFinite(options.requiredSupport) ||
        options.requiredSupport < 0 ||
        options.requiredSupport > 1)
    )
      throw new Error('Evidence support threshold must be between zero and one.');
  }
  state(): EvidenceState {
    return structuredClone(this.current);
  }
  install(hooks: RuntimeHooks): () => void {
    const off = [
      hooks.register({
        id: 'core.evidence/receive',
        on: 'message:received',
        mode: 'guard',
        timeoutMs: 10000,
        handle: async ({ mode, text, dataStoreSpecs }, ctx) => {
          this.current = {
            status: 'idle',
            sources: [],
            message: 'No additional evidence search requested.',
          };
          if (!['chat', 'command'].includes(mode) || !dataStoreSpecs?.length) return;
          if (dataStoreSpecs.length > 8 || text.length > 32768)
            return { kind: 'block', reason: 'Evidence request exceeds the scope budget.' };
          const scopes = dataStoreSpecs.map((s) => ({
            dataStore: s.dataStore,
            ...(s.filter ? { filter: s.filter } : {}),
          }));
          if (
            scopes.some(
              (s) =>
                !/^projects\/[^/]+\/locations\/[^/]+\/collections\/[^/]+\/dataStores\/[^/]+$/.test(
                  s.dataStore,
                ),
            )
          )
            return {
              kind: 'block',
              reason: 'Select valid, authorized data stores before searching.',
            };
          this.current = {
            status: 'searching',
            sources: [],
            message: `Searching ${scopes.length} selected data store(s).`,
          };
          try {
            const result = await this.options.search.search({
              query: text,
              dataStoreSpecs: scopes,
              pageSize: 24,
              snippets: true,
              signal: ctx.signal,
            });
            ctx.signal.throwIfAborted();
            // Keep only explicitly selected store resources. Never trust a response to broaden scope.
            const dedup = new Map<string, SearchHit>();
            for (const hit of result.results.slice(0, 48)) {
              if (
                !scopes.some(
                  (s) =>
                    hit.documentName.startsWith(`${s.dataStore}/branches/`) ||
                    hit.documentName.startsWith(`${s.dataStore}/documents/`),
                )
              )
                continue;
              if (!hit.snippet?.trim() || dedup.has(hit.documentName)) continue;
              dedup.set(hit.documentName, {
                ...hit,
                id: `e${dedup.size + 1}`,
                title: hit.title?.slice(0, 256),
                snippet: hit.snippet.slice(0, 4000),
              });
            }
            let candidates = [...dedup.values()].slice(0, 24);
            let ranked = false;
            if (candidates.length > 1) {
              try {
                const order = await this.options.rank.rank(
                  text,
                  candidates.map((h) => ({ id: h.id, title: h.title, content: h.snippet })),
                  { topN: 8, signal: ctx.signal },
                );
                ctx.signal.throwIfAborted();
                const byId = new Map(candidates.map((h) => [h.id, h]));
                const selected: SearchHit[] = [];
                for (const entry of order) {
                  const hit = byId.get(entry.id);
                  if (hit) {
                    selected.push(hit);
                    byId.delete(entry.id);
                  }
                }
                if (selected.length) {
                  candidates = [...selected, ...byId.values()];
                  ranked = true;
                }
              } catch {
                ctx.signal.throwIfAborted();
              }
            }
            candidates = candidates.slice(0, 8);
            const facts = candidates.map((h) => ({
              text: h.snippet!,
              attributes: {
                source: h.documentName,
                title: h.title ?? h.id,
                ...(safeUri(h.uri) ? { uri: safeUri(h.uri)! } : {}),
              },
            }));
            this.tasks.set(ctx.taskId, facts);
            this.current = {
              status: 'ready',
              sources: candidates.map((h) => ({
                id: h.documentName,
                title: h.title ?? h.id,
                ...(safeUri(h.uri) ? { uri: safeUri(h.uri)! } : {}),
              })),
              message: `${facts.length} source excerpt(s) assembled${ranked ? ' and reranked' : ''}. Excerpts may be incomplete.`,
            };
            if (!facts.length && this.options.requiredSupport !== undefined)
              return {
                kind: 'block',
                reason: 'No evidence excerpts were available in the selected scope.',
              };
            const entries: ResolvedContext[] = facts.map((f, i) => ({
              ref: {
                id: `evidence:${ctx.taskId}:${i}`,
                kind: 'brief',
                surface: ctx.surface,
                title: f.attributes.title,
                live: false,
              },
              value: {
                as: 'text',
                text: JSON.stringify({ source: f.attributes, excerpt: f.text }),
                mimeType: 'application/json',
              },
            }));
            return { kind: 'context', entries };
          } catch (error) {
            ctx.signal.throwIfAborted();
            this.current = {
              status: 'unavailable',
              sources: [],
              message: 'Evidence retrieval is unavailable. No additional sources were attached.',
            };
            if (this.options.requiredSupport !== undefined)
              return { kind: 'block', reason: this.current.message };
            return;
          }
        },
      }),
      hooks.register({
        id: 'core.evidence/check',
        on: 'model:response',
        mode: 'guard',
        timeoutMs: 10000,
        handle: async ({ text }, ctx) => {
          const facts = this.tasks.get(ctx.taskId);
          if (!facts?.length || !text.trim()) return;
          // Executable command blocks are checked by the plan/host verifier, not a text support score.
          if (/```(?:cmd|plan)\b/.test(text)) return;
          if (text.length > 32768) {
            this.current = {
              ...this.current,
              status: 'unavailable',
              message: 'Answer exceeds the claim-check budget.',
            };
            return this.options.requiredSupport === undefined
              ? undefined
              : { kind: 'block', reason: this.current.message };
          }
          this.current = {
            ...this.current,
            status: 'checking',
            message: 'Checking claims against the selected source excerpts.',
          };
          try {
            const result = await this.options.grounding.check(text, facts, {
              claimLevelScore: true,
              signal: ctx.signal,
            });
            ctx.signal.throwIfAborted();
            if (
              !result.claims?.length ||
              !Number.isFinite(result.supportScore) ||
              result.supportScore < 0 ||
              result.supportScore > 1
            )
              throw new Error('No checkable claims returned.');
            this.current = {
              ...this.current,
              status: 'checked',
              score: result.supportScore,
              checkedClaims: result.claims.length,
              message:
                'Support score applies to these excerpts. It does not verify calculations or completeness.',
            };
            if (
              this.options.requiredSupport !== undefined &&
              result.supportScore < this.options.requiredSupport
            )
              return {
                kind: 'block',
                reason: 'The answer did not meet the configured evidence support threshold.',
              };
          } catch {
            ctx.signal.throwIfAborted();
            this.current = {
              ...this.current,
              status: 'unavailable',
              message: 'Claim checking was unavailable; the answer is not verified.',
            };
            if (this.options.requiredSupport !== undefined)
              return { kind: 'block', reason: this.current.message };
          }
          return;
        },
      }),
      hooks.register({
        id: 'core.evidence/cleanup',
        on: 'task:finished',
        mode: 'observe',
        handle: (_payload, ctx) => {
          this.tasks.delete(ctx.taskId);
          if (['searching', 'checking'].includes(this.current.status))
            this.current = {
              ...this.current,
              status: 'unavailable',
              message: 'Evidence processing did not complete.',
            };
        },
      }),
    ];
    return () => {
      off.forEach((dispose) => dispose());
      this.tasks.clear();
    };
  }
}
