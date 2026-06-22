import type { ContextKind, UnitDescriptor } from '@ge/contracts';
import { asSessionId } from '@ge/contracts';
import type { AssistantPath, WifConfig } from '@ge/gemini-client';
import { StreamAssistClient, WifTokenClient } from '@ge/gemini-client';
import { AssistSession } from '@ge/runtime';
import type { AuthClient, DocBridge } from '@ge/runtime';
import type { TriggerRegistry } from '@ge/triggers';

/**
 * Everything the shell needs to reach Gemini Enterprise as the signed-in user.
 * No Google credential here — only the WIF pool/provider the user's Entra token federates to.
 */
export interface ShellConfig {
  assistant: AssistantPath;
  wif: WifConfig;
  modelId?: string;
  /** Optional audited/CORS egress proxy (ADR-0001); omitted → call Discovery Engine directly. */
  proxyUrl?: string;
}

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
}

export interface ComposedSession {
  session: AssistSession;
  tokens: WifTokenClient;
  client: StreamAssistClient;
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
      identity: identity.username,
      ...(config.modelId ? { modelId: config.modelId } : {}),
      ...(config.proxyUrl ? { proxyUrl: config.proxyUrl } : {}),
    },
    opts.fetchImpl,
  );

  const session = new AssistSession(bridge, client, {
    unit,
    ...(opts.autoAttach ? { autoAttach: opts.autoAttach } : {}),
    ...(opts.triggers ? { triggers: opts.triggers } : {}),
    ...(opts.resumeSessionId ? { resumeSessionId: asSessionId(opts.resumeSessionId) } : {}),
  });

  return { session, tokens, client };
}
