import type { ContextKind, ReleaseProfileName, UnitDescriptor } from '@ge/contracts';
import { asSessionId, filterManifestForReleaseProfile } from '@ge/contracts';
import type {
  AssistantPath,
  GeminiSkillMention,
  GeminiWidgetConfig,
  WifConfig,
} from '@ge/gemini-client';
import {
  StreamAssistClient,
  WifTokenClient,
  ensureSkillAgent,
  listAvailableAgentViews,
  listEngineDataStores,
} from '@ge/gemini-client';
import type {
  AgentView,
  EngineDataStore,
  EnsureSkillInput,
  EnsureSkillResult,
} from '@ge/gemini-client';
import { AssistSession } from '@ge/runtime';
import type { AuthClient, DocBridge } from '@ge/runtime';
import { GraphClient, GraphSharedStore } from '@ge/graph-client';
import type { TriggerRegistry } from '@ge/triggers';

/**
 * Everything the shell needs to reach Gemini Enterprise as the signed-in user.
 * No Google credential here — only the WIF pool/provider the user's Entra token federates to.
 */
export interface ShellConfig {
  assistant: AssistantPath;
  widget?: GeminiWidgetConfig;
  wif: WifConfig;
  modelId?: string;
  /** Gemini Enterprise skills/agents mounted on ordinary chat turns. */
  skills?: string[];
  /** GE widget-style mentions for ordinary chat skills. */
  skillMentions?: GeminiSkillMention[];
  /** Planner skill used only for the plan-confirm route. */
  plannerSkills?: string[];
  plannerSkillMentions?: GeminiSkillMention[];
  /** Commander skill used only for the constrained command-loop route. */
  commandSkills?: string[];
  commandSkillMentions?: GeminiSkillMention[];
  /** Optional audited/CORS egress proxy (ADR-0001); omitted → call Discovery Engine directly. */
  proxyUrl?: string;
  /** Release profile that narrows capabilities before UI/model/executor exposure. */
  releaseProfile?: ReleaseProfileName;
}

/**
 * Skill bundle reference files, bundled at build time (not fetched — there is no server to fetch
 * them from in this client-direct architecture) and exposed read-only at `/skills` in DocFs, so
 * the client-side command loop can `ls`/`cat`/`find`/`grep` a specific reference file the same way
 * it does `/doc`/`/work`, instead of relying solely on whatever `skillsSpec`'s opaque server-side
 * name reference (see `stream-assist.ts`) causes the engine to inject. Keys are paths relative to
 * `skill/`, e.g. `"m365-surface-commander/references/excel-semantics.md"`.
 */
export const SKILL_FILES: Record<string, string> = Object.fromEntries(
  Object.entries(
    import.meta.glob('../../../skill/**/*.md', { query: '?raw', import: 'default', eager: true }),
  ).map(([path, content]) => [path.replace(/^.*\/skill\//, ''), content as string]),
);

export interface ComposeOptions {
  config: ShellConfig;
  auth: AuthClient;
  bridge: DocBridge;
  unit: UnitDescriptor;
  triggers?: TriggerRegistry;
  autoAttach?: ContextKind[];
  /** Resume a prior session (persisted in host metadata) so constructed context survives. */
  resumeSessionId?: string;
  fetchImpl?: typeof fetch;
  /**
   * Boot-time skill provisioning ("warm-up"): the add-in ensures its own skill agents exist and
   * match, using the signed-in user's WIF token — no out-of-band admin/widget step (ADR-0001).
   * Idempotent and best-effort: a cheap GetAgent per skill, writing only on drift; a failure is
   * captured in `ComposedSession.warmUp`, never thrown, so provisioning can't block chat.
   */
  warmUpSkills?: EnsureSkillInput[];
  /** Billing/quota project for warm-up calls when the WIF token doesn't carry one. */
  warmUpQuotaProject?: string;
  /**
   * Discover, at boot, the agents/skills available to the user (`:listAvailableAgentViews`) and the
   * federated data stores attached to the engine (`engines.get` → `dataStoreIds`), for the skill/
   * data-store pickers. WIF-authenticated, no widget, no admin grant. Best-effort — failures are
   * swallowed and the corresponding list is empty; never blocks the session.
   */
  discoverCatalog?: boolean;
}

export interface ComposedSession {
  session: AssistSession;
  tokens: WifTokenClient;
  client: StreamAssistClient;
  /** Result of boot-time skill provisioning, one entry per `warmUpSkills` input (empty if none). */
  warmUp: Array<EnsureSkillResult | { action: 'error'; agentId: string; error: string }>;
  /** Agents/skills available to the user (empty unless `discoverCatalog`). Source for the skill picker. */
  availableAgents: AgentView[];
  /** Federated data stores on the engine (empty unless `discoverCatalog`). Source for the data-store picker. */
  availableDataStores: EngineDataStore[];
}

/**
 * Wire the chain once: user identity → WIF token exchange → Discovery Engine client →
 * a ready `AssistSession` bound to this surface's bridge. The identity is resolved up front
 * so every turn's provenance is stamped with the signed-in user.
 */
export async function composeSession(opts: ComposeOptions): Promise<ComposedSession> {
  const { config, auth, bridge, unit } = opts;
  const identity = await auth.getIdentity();

  const tokens = new WifTokenClient(auth, config.wif, opts.fetchImpl);
  const client = new StreamAssistClient(
    tokens,
    {
      assistant: config.assistant,
      ...(config.widget ? { widget: config.widget } : {}),
      identity: identity.username,
      ...(config.modelId ? { modelId: config.modelId } : {}),
      ...(config.skills?.length ? { skills: config.skills } : {}),
      ...(config.skillMentions?.length ? { skillMentions: config.skillMentions } : {}),
      ...(config.plannerSkills?.length ? { plannerSkills: config.plannerSkills } : {}),
      ...(config.plannerSkillMentions?.length
        ? { plannerSkillMentions: config.plannerSkillMentions }
        : {}),
      ...(config.commandSkills?.length ? { commandSkills: config.commandSkills } : {}),
      ...(config.commandSkillMentions?.length
        ? { commandSkillMentions: config.commandSkillMentions }
        : {}),
      ...(config.proxyUrl ? { proxyUrl: config.proxyUrl } : {}),
    },
    opts.fetchImpl,
  );

  // Warm-up: self-provision skill agents as the signed-in user (best-effort, idempotent).
  const warmUp: ComposedSession['warmUp'] = [];
  for (const skill of opts.warmUpSkills ?? []) {
    try {
      warmUp.push(
        await ensureSkillAgent({ assistant: config.assistant }, skill, {
          tokens,
          ...(opts.fetchImpl ? { fetchImpl: opts.fetchImpl } : {}),
          ...(opts.warmUpQuotaProject ? { quotaProject: opts.warmUpQuotaProject } : {}),
        }),
      );
    } catch (err) {
      warmUp.push({
        action: 'error',
        agentId: skill.agentId,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  // Catalog discovery (best-effort): what the user can mount (skills) and ground on (data stores).
  let availableAgents: AgentView[] = [];
  let availableDataStores: EngineDataStore[] = [];
  if (opts.discoverCatalog) {
    const dsOpts = {
      tokens,
      ...(opts.fetchImpl ? { fetchImpl: opts.fetchImpl } : {}),
      ...(opts.warmUpQuotaProject ? { quotaProject: opts.warmUpQuotaProject } : {}),
    };
    const cfg = { assistant: config.assistant };
    const [agentsRes, storesRes] = await Promise.allSettled([
      listAvailableAgentViews(cfg, dsOpts),
      listEngineDataStores(cfg, { ...dsOpts, enrich: false }),
    ]);
    if (agentsRes.status === 'fulfilled') availableAgents = agentsRes.value;
    if (storesRes.status === 'fulfilled') availableDataStores = storesRes.value;
  }

  // `/shared` cross-surface handoff (see docs/ACCESS-MODEL.md Plane B): only wired when this
  // AuthClient actually carries a Graph token source. Absent that, `share`/`/shared` degrade to
  // the runtime's own "not configured" behavior — never a hard dependency on Graph consent.
  // `estateWritesEnabled` mirrors that same condition: `share` is live whenever there's a real
  // store to write to, gated per-use by the UI's `ShareApprovalCard` (approveShare) — the same
  // fail-closed human-in-the-loop every other write in this app already requires.
  const sharedStore = auth.getGraphToken
    ? new GraphSharedStore(
        new GraphClient(
          { getGraphToken: (scopes) => auth.getGraphToken!(scopes) },
          {},
          opts.fetchImpl,
        ),
      )
    : undefined;

  const session = new AssistSession(bridge, client, {
    unit,
    skillFiles: SKILL_FILES,
    ...(sharedStore ? { sharedStore, estateWritesEnabled: true } : {}),
    ...(opts.autoAttach ? { autoAttach: opts.autoAttach } : {}),
    ...(opts.triggers ? { triggers: opts.triggers } : {}),
    ...(opts.resumeSessionId ? { resumeSessionId: asSessionId(opts.resumeSessionId) } : {}),
    ...(config.releaseProfile
      ? {
          capabilityFilter: (manifest) =>
            filterManifestForReleaseProfile(manifest, config.releaseProfile!),
        }
      : {}),
  });

  return { session, tokens, client, warmUp, availableAgents, availableDataStores };
}
