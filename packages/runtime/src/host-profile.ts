import {
  CAPABILITY_REGISTRY,
  type ActuationKind,
  type CapabilityManifest,
  type HostCapabilityProfile,
  type SupportDecision,
  type Surface,
} from '@ge/contracts';

/** Injectable narrow adapter: no Office.js dependency is needed in simulators or offline tooling. */
export interface HostEnvironment {
  platform?: string;
  build?: string;
  mode?: HostCapabilityProfile['mode'];
  isSetSupported(name: string, version: string): boolean;
}
export function hostCapabilityProfile(
  surface: Surface,
  environment: HostEnvironment,
): HostCapabilityProfile {
  const requirements: Record<string, boolean> = {};
  for (const entry of CAPABILITY_REGISTRY.filter((entry) => entry.surface === surface)) {
    for (const set of entry.requirementSets) {
      const key = `${set.name}:${set.minVersion}`;
      if (Object.hasOwn(requirements, key)) continue;
      try {
        requirements[key] = environment.isSetSupported(set.name, set.minVersion) === true;
      } catch {
        requirements[key] = false;
      }
    }
  }
  return {
    surface,
    platform: (environment.platform ?? 'unknown').slice(0, 64),
    build: (environment.build ?? 'unknown').slice(0, 64),
    mode: environment.mode ?? 'unknown',
    requirements,
  };
}
const composeOnly = new Set<ActuationKind>([
  'set-body',
  'set-subject',
  'set-recipients',
  'add-attachment',
]);
export function capabilitySupport(
  surface: Surface,
  kind: ActuationKind,
  profile: HostCapabilityProfile,
): SupportDecision {
  if (surface !== profile.surface) return { supported: false, reason: 'wrong_surface' };
  if (surface !== 'outlook' && profile.mode === 'read')
    return { supported: false, reason: 'read_only' };
  if (surface === 'outlook' && composeOnly.has(kind) && profile.mode !== 'compose')
    return { supported: false, reason: 'compose_required' };
  const entry = CAPABILITY_REGISTRY.find(
    (entry) => entry.kind === kind && entry.surface === surface,
  );
  // An uncatalogued member has no verified support contract. Offline ports can omit a profile.
  if (!entry) return { supported: false, reason: 'unsupported_api', missing: ['support-contract'] };
  const missing = entry.requirementSets
    .map((set) => `${set.name}:${set.minVersion}`)
    .filter((key) => profile.requirements[key] !== true);
  return missing.length
    ? { supported: false, reason: 'unsupported_api', missing }
    : { supported: true };
}
export function negotiateCapabilities(
  manifest: CapabilityManifest,
  profile?: HostCapabilityProfile,
): CapabilityManifest {
  return {
    ...manifest,
    contextKinds: [...manifest.contextKinds],
    ...(manifest.reads ? { reads: [...manifest.reads] } : {}),
    actuations: manifest.actuations
      .filter(
        (operation) =>
          !profile || capabilitySupport(manifest.surface, operation.kind, profile).supported,
      )
      .map((operation) => ({ ...operation })),
  };
}
