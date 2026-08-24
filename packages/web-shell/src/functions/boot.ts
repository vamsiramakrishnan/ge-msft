import type { AuthClient } from '@ge/runtime';
import { StreamAssistClient, WifTokenClient } from '@ge/gemini-client';
import { NaaAuthClient } from '../auth-client.js';
import {
  authOptionsFromEnv,
  msalConfigFromEnv,
  shellConfigFromEnv,
  type RawEnv,
} from '../taskpane/config.js';
import { createMsal } from '../taskpane/msal.js';
import { createGeAskAssist } from './ge-ask-assist.js';
import type { GeAskAssist } from './ge-ask.js';

/**
 * Functions-runtime boot: the same client-direct chain the task pane composes in
 * `composeSession` (identity → WIF exchange → Discovery Engine), minus the bridge/session —
 * `=GE.ASK` is a pure grounded ask. No Google credential exists anywhere on this path; only the
 * signed-in user's short-lived Entra token federates.
 */
export async function createDefaultGeAskAssist(env: RawEnv): Promise<GeAskAssist> {
  const config = shellConfigFromEnv(env);
  const msal = await createMsal({ ...msalConfigFromEnv(env) });
  const auth: AuthClient = new NaaAuthClient(msal, authOptionsFromEnv(env));
  const identity = await auth.getIdentity();
  const tokens = new WifTokenClient(auth, config.wif);
  const client = new StreamAssistClient(tokens, {
    assistant: config.assistant,
    ...(config.widget ? { widget: config.widget } : {}),
    identity: identity.username,
    ...(config.modelId ? { modelId: config.modelId } : {}),
    ...(config.proxyUrl ? { proxyUrl: config.proxyUrl } : {}),
  });
  return createGeAskAssist(client);
}
