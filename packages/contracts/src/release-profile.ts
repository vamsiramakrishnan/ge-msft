import { z } from 'zod';
import {
  ActuationKindSchema,
  CapabilityManifestSchema,
  type Actuation,
  type ActuationKind,
  type CapabilityManifest,
} from './capability.js';
import { SurfaceSchema, type Surface } from './context.js';

export const ReleaseProfileNameSchema = z.enum(['development', 'internal-alpha-word-excel']);
export type ReleaseProfileName = z.infer<typeof ReleaseProfileNameSchema>;

export const CapabilityAvailabilitySchema = z.discriminatedUnion('state', [
  z.object({ state: z.literal('supported') }),
  z.object({ state: z.literal('degraded'), reason: z.string() }),
  z.object({
    state: z.literal('unavailable'),
    reason: z.string(),
    remediation: z.string().optional(),
  }),
]);
export type CapabilityAvailability = z.infer<typeof CapabilityAvailabilitySchema>;

export const ReleaseProfileSchema = z.object({
  name: ReleaseProfileNameSchema,
  enabledSurfaces: z.array(SurfaceSchema),
  disabledSurfaces: z.array(SurfaceSchema),
  enabledActuations: z.record(SurfaceSchema, z.array(ActuationKindSchema)),
  durableProvenanceRequired: z.boolean(),
  crossSurfacePlans: z.literal(false),
  /**
   * Whether this profile permits `share`/`/shared` (the cross-surface Graph-backed handoff
   * store — the first estate write, see `AssistSessionOptions.estateWritesEnabled`).
   * `packages/web-shell/src/compose.ts` reads this directly: `estateWritesEnabled` is `true`
   * only when BOTH a `sharedStore` is wired AND this flag is `true` for the active profile —
   * so this is a real, enforced tenant/deployment lever, not decorative.
   */
  estateWrites: z.boolean(),
  arbitraryCodeExecution: z.literal(false),
});
export type ReleaseProfile = z.infer<typeof ReleaseProfileSchema>;

const ALL_SURFACES = SurfaceSchema.options;

export const INTERNAL_ALPHA_WORD_EXCEL_PROFILE: ReleaseProfile = {
  name: 'internal-alpha-word-excel',
  enabledSurfaces: ['word', 'excel'],
  disabledSurfaces: ALL_SURFACES.filter((s) => s !== 'word' && s !== 'excel'),
  enabledActuations: {
    word: ['tracked-change', 'add-comment', 'comment-reply'],
    excel: ['write-cells', 'format-cells', 'add-comment', 'comment-reply'],
    powerpoint: [],
    onenote: [],
    outlook: [],
    teams: [],
  },
  durableProvenanceRequired: true,
  crossSurfacePlans: false,
  estateWrites: true,
  arbitraryCodeExecution: false,
};

export const DEVELOPMENT_PROFILE: ReleaseProfile = {
  name: 'development',
  enabledSurfaces: [...ALL_SURFACES],
  disabledSurfaces: [],
  enabledActuations: {
    word: [...ActuationKindSchema.options],
    excel: [...ActuationKindSchema.options],
    powerpoint: [...ActuationKindSchema.options],
    onenote: [...ActuationKindSchema.options],
    outlook: [...ActuationKindSchema.options],
    teams: [...ActuationKindSchema.options],
  },
  durableProvenanceRequired: false,
  crossSurfacePlans: false,
  estateWrites: true,
  arbitraryCodeExecution: false,
};

export function releaseProfile(name: ReleaseProfileName): ReleaseProfile {
  switch (name) {
    case 'internal-alpha-word-excel':
      return INTERNAL_ALPHA_WORD_EXCEL_PROFILE;
    case 'development':
      return DEVELOPMENT_PROFILE;
  }
}

export function isSurfaceEnabledForProfile(surface: Surface, profile: ReleaseProfile): boolean {
  return profile.enabledSurfaces.includes(surface);
}

export function enabledActuationKindsForProfile(
  surface: Surface,
  profile: ReleaseProfile,
): ReadonlySet<ActuationKind> {
  return new Set(profile.enabledActuations[surface] ?? []);
}

export function filterManifestForReleaseProfile(
  manifest: CapabilityManifest,
  profile: ReleaseProfile | ReleaseProfileName,
): CapabilityManifest {
  const parsed = CapabilityManifestSchema.parse(manifest);
  const p = typeof profile === 'string' ? releaseProfile(profile) : profile;
  if (!isSurfaceEnabledForProfile(parsed.surface, p)) {
    return { ...parsed, contextKinds: [], reads: [], actuations: [] };
  }
  const allowed = enabledActuationKindsForProfile(parsed.surface, p);
  const actuations = parsed.actuations.filter((a) => allowed.has(a.kind));
  return { ...parsed, actuations };
}

export type RequirementSetProbe = Record<string, string | undefined>;

export interface CapabilityAvailabilityInput {
  profile: ReleaseProfile | ReleaseProfileName;
  manifest: CapabilityManifest;
  kind: ActuationKind;
  requirementSets?: RequirementSetProbe;
  activeSurface: Surface;
  platform?: string;
  documentMode?: 'read-write' | 'read-only' | 'unknown';
  grantedScopes?: readonly string[];
  tenantPolicy?: 'allow' | 'deny' | 'unknown';
  featureFlags?: Readonly<Record<string, boolean | undefined>>;
  durableProvenance?: 'supported' | 'unsupported' | 'unknown';
}

const WRITE_REQUIREMENTS: Partial<
  Record<Surface, Partial<Record<ActuationKind, RequirementSetProbe>>>
> = {
  word: {
    'tracked-change': { WordApi: '1.4' },
    'add-comment': { WordApi: '1.4' },
    'comment-reply': { WordApi: '1.4' },
  },
  excel: {
    'write-cells': { ExcelApi: '1.1' },
    'format-cells': { ExcelApi: '1.1' },
    'add-comment': { ExcelApi: '1.10' },
    'comment-reply': { ExcelApi: '1.10' },
  },
};

export function capabilityAvailability(input: CapabilityAvailabilityInput): CapabilityAvailability {
  const profile = typeof input.profile === 'string' ? releaseProfile(input.profile) : input.profile;
  const manifest = CapabilityManifestSchema.parse(input.manifest);

  if (manifest.surface !== input.activeSurface) {
    return {
      state: 'unavailable',
      reason: `capability belongs to ${manifest.surface}, active host is ${input.activeSurface}`,
    };
  }
  if (!isSurfaceEnabledForProfile(input.activeSurface, profile)) {
    return {
      state: 'unavailable',
      reason: `${input.activeSurface} is disabled by release profile ${profile.name}`,
    };
  }
  const profileKinds = enabledActuationKindsForProfile(input.activeSurface, profile);
  if (!profileKinds.has(input.kind)) {
    return {
      state: 'unavailable',
      reason: `${input.kind} is not enabled by release profile ${profile.name}`,
    };
  }
  if (!manifest.actuations.some((a: Actuation) => a.kind === input.kind)) {
    return { state: 'unavailable', reason: `${input.kind} is not statically closed` };
  }
  if (input.documentMode !== undefined && input.documentMode !== 'read-write') {
    return {
      state: 'unavailable',
      reason: 'the current document is not writable',
      remediation: 'Open an editable document or workbook and retry.',
    };
  }
  if (input.tenantPolicy !== undefined && input.tenantPolicy !== 'allow') {
    return {
      state: 'unavailable',
      reason: 'tenant policy does not allow this mutation',
      remediation: 'Ask the tenant administrator to enable the alpha policy for this capability.',
    };
  }
  if (profile.durableProvenanceRequired && input.durableProvenance !== 'supported') {
    return {
      state: 'unavailable',
      reason: 'durable provenance is required for alpha writes',
      remediation: 'Use a host/platform where provenance persistence is certified.',
    };
  }

  const required = WRITE_REQUIREMENTS[input.activeSurface]?.[input.kind] ?? {};
  for (const [setName, minVersion] of Object.entries(required)) {
    const actual = input.requirementSets?.[setName];
    if (!actual) {
      return {
        state: 'unavailable',
        reason: `unknown ${setName} requirement-set state`,
        remediation: `Probe ${setName} ${minVersion} or disable this write.`,
      };
    }
    if (!versionAtLeast(actual, minVersion)) {
      return {
        state: 'unavailable',
        reason: `${setName} ${actual} is below required ${minVersion}`,
      };
    }
  }

  return { state: 'supported' };
}

function versionAtLeast(actual: string, minimum: string | undefined): boolean {
  if (!minimum) return true;
  const a = actual.split('.').map((n) => Number(n));
  const m = minimum.split('.').map((n) => Number(n));
  for (let i = 0; i < Math.max(a.length, m.length); i++) {
    const av = a[i] ?? 0;
    const mv = m[i] ?? 0;
    if (av > mv) return true;
    if (av < mv) return false;
  }
  return true;
}
