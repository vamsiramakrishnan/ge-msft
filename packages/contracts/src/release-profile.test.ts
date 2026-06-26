import { describe, expect, it } from 'vitest';
import {
  capabilityAvailability,
  filterManifestForReleaseProfile,
  INTERNAL_ALPHA_WORD_EXCEL_PROFILE,
} from './release-profile.js';
import type { CapabilityManifest } from './capability.js';

const excelManifest: CapabilityManifest = {
  surface: 'excel',
  contextKinds: ['range'],
  reads: ['outline', 'read', 'search'],
  actuations: [
    {
      kind: 'write-cells',
      surface: 'excel',
      title: 'Write cells',
      reversible: true,
    },
    {
      kind: 'insert-chart',
      surface: 'excel',
      title: 'Insert chart',
      reversible: true,
    },
    {
      kind: 'format-cells',
      surface: 'excel',
      title: 'Format cells',
      reversible: true,
    },
  ],
};

describe('internal-alpha-word-excel release profile', () => {
  it('enables only Word and Excel surfaces', () => {
    expect(INTERNAL_ALPHA_WORD_EXCEL_PROFILE.enabledSurfaces).toEqual(['word', 'excel']);
    expect(INTERNAL_ALPHA_WORD_EXCEL_PROFILE.disabledSurfaces).toEqual([
      'powerpoint',
      'onenote',
      'outlook',
      'teams',
    ]);
  });

  it('filters Excel to alpha-approved write classes only', () => {
    const filtered = filterManifestForReleaseProfile(
      excelManifest,
      INTERNAL_ALPHA_WORD_EXCEL_PROFILE,
    );
    expect(filtered.actuations.map((a) => a.kind)).toEqual(['write-cells', 'format-cells']);
  });

  it('removes all writes for disabled surfaces', () => {
    const filtered = filterManifestForReleaseProfile(
      { ...excelManifest, surface: 'powerpoint' },
      INTERNAL_ALPHA_WORD_EXCEL_PROFILE,
    );
    expect(filtered.actuations).toEqual([]);
    expect(filtered.reads).toEqual([]);
  });

  it('fails closed when durable provenance is unknown for alpha writes', () => {
    expect(
      capabilityAvailability({
        profile: INTERNAL_ALPHA_WORD_EXCEL_PROFILE,
        manifest: excelManifest,
        kind: 'write-cells',
        activeSurface: 'excel',
        requirementSets: { ExcelApi: '1.12' },
        documentMode: 'read-write',
        tenantPolicy: 'allow',
        durableProvenance: 'unknown',
      }),
    ).toMatchObject({ state: 'unavailable', reason: expect.stringMatching(/provenance/i) });
  });

  it('fails closed when requirement-set probes are missing', () => {
    expect(
      capabilityAvailability({
        profile: INTERNAL_ALPHA_WORD_EXCEL_PROFILE,
        manifest: excelManifest,
        kind: 'write-cells',
        activeSurface: 'excel',
        documentMode: 'read-write',
        tenantPolicy: 'allow',
        durableProvenance: 'supported',
      }),
    ).toMatchObject({ state: 'unavailable', reason: expect.stringMatching(/unknown ExcelApi/i) });
  });

  it('supports an alpha write only after profile, static closure, host, policy, and provenance pass', () => {
    expect(
      capabilityAvailability({
        profile: INTERNAL_ALPHA_WORD_EXCEL_PROFILE,
        manifest: excelManifest,
        kind: 'write-cells',
        activeSurface: 'excel',
        requirementSets: { ExcelApi: '1.12' },
        documentMode: 'read-write',
        tenantPolicy: 'allow',
        durableProvenance: 'supported',
      }),
    ).toEqual({ state: 'supported' });
  });
});
