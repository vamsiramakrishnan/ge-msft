import { existsSync, readFileSync, statSync } from 'node:fs';
import { z } from 'zod';
import type { AssistantPath } from '@ge/gemini-client';
import { loadShellEnv } from './live-env.js';

export interface TenantWorkflowProvider {
  assistant: AssistantPath;
  commander: { resource: string; label: string };
  getAccessToken(): Promise<string>;
  fetchImpl?: typeof fetch;
}
const identifier = z.string().regex(/^[A-Za-z0-9][A-Za-z0-9._:-]{1,127}$/);
const location = z.enum(['global', 'us', 'eu', 'asia-northeast1', 'australia-southeast1']);
const skill = z
  .string()
  .regex(
    /^(?:[A-Za-z0-9][A-Za-z0-9._-]{0,127}=)?(?:projects\/[A-Za-z0-9._:-]+\/locations\/[A-Za-z0-9._-]+\/(?:collections\/[A-Za-z0-9._-]+\/engines\/[A-Za-z0-9._:-]+\/)?assistants\/[A-Za-z0-9._:-]+\/agents\/)?[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/,
  );
/** Explicit token configuration only. This does not attempt login, gcloud or implicit global-region fallback. */
export function configuredTenantWorkflowProvider(
  environment: Record<string, string | undefined> = { ...loadShellEnv(), ...process.env },
): TenantWorkflowProvider | undefined {
  const project =
    environment.GE_PROJECT_NUMBER ?? environment.GE_PROJECT ?? environment.VITE_GCP_PROJECT;
  const engine = environment.GE_ENGINE ?? environment.VITE_GE_ENGINE;
  const region = environment.GE_LOCATION ?? environment.VITE_GCP_LOCATION;
  const command =
    environment.GE_SURFACE_COMMANDER_SKILL ?? environment.VITE_GE_SURFACE_COMMANDER_SKILL;
  const tokenFile =
    (environment.GE_WIF_ACCESS_TOKEN_FILE ?? environment.GE_ACCESS_TOKEN_FILE)?.trim() || undefined;
  const inline =
    (environment.GE_WIF_ACCESS_TOKEN ?? environment.GE_ACCESS_TOKEN)?.trim() || undefined;
  if (
    !identifier.safeParse(project).success ||
    !identifier.safeParse(engine).success ||
    !location.safeParse(region).success ||
    !skill.safeParse(command).success
  )
    return undefined;
  if (!inline && (!tokenFile || !existsSync(tokenFile))) return undefined;
  const [label, resource] = command!.includes('=')
    ? command!.split('=', 2)
    : [command!.split('/').at(-1)!, command!];
  return {
    assistant: { project: project!, location: region!, engine: engine! },
    commander: { label: label!, resource: resource! },
    async getAccessToken() {
      if (!inline && tokenFile && statSync(tokenFile).size > 16384)
        throw new Error('Configured token exceeds its size budget.');
      const token = inline ?? (tokenFile ? readFileSync(tokenFile, 'utf8').trim() : '');
      if (!token || token.length > 16384 || /[\r\n]/.test(token))
        throw new Error('A current explicitly configured access token is required.');
      return token;
    },
  };
}
