import { describe, expect, it } from 'vitest';
import { ActuationKindSchema } from './capability.js';
import {
  assertCapabilityRegistryConsistent,
  capabilityRegistryEntries,
  registryEntriesForSurface,
  registryEntryForKindAndSurface,
} from './capability-registry.js';
import { ContextKindSchema, SurfaceSchema } from './context.js';

describe('capability registry', () => {
  it('is internally consistent and schema-valid', () => {
    expect(() => assertCapabilityRegistryConsistent()).not.toThrow();
  });

  it('references only real surfaces, actuation kinds, and context kinds', () => {
    const kinds = new Set(ActuationKindSchema.options);
    const surfaces = new Set(SurfaceSchema.options);
    const contextKinds = new Set(ContextKindSchema.options);
    for (const entry of capabilityRegistryEntries()) {
      expect(kinds.has(entry.kind)).toBe(true);
      expect(surfaces.has(entry.surface)).toBe(true);
      for (const contextKind of entry.contextKinds) {
        expect(contextKinds.has(contextKind)).toBe(true);
      }
    }
  });

  it('keeps core verbs and specialized slash commands separate', () => {
    for (const entry of capabilityRegistryEntries()) {
      if (entry.exposure === 'core-verb') {
        expect(entry.command.startsWith('/')).toBe(false);
      }
      if (entry.exposure === 'specialized') {
        expect(entry.command).toBe(`/${entry.kind}`);
      }
    }
  });

  it('tracks the first Excel advanced Office.js promotions without making them live authority', () => {
    const excelKinds = registryEntriesForSurface('excel').map((entry) => entry.kind);
    expect(excelKinds).toEqual(
      expect.arrayContaining([
        'insert-pivot',
        'sort-range',
        'filter-range',
        'manage-worksheet',
        'format-chart',
      ]),
    );
    expect(registryEntryForKindAndSurface('insert-pivot', 'excel')).toMatchObject({
      status: 'promotable',
      command: '/insert-pivot',
      requirementSets: [{ name: 'ExcelApi', minVersion: '1.8' }],
    });
  });

  it('teaches Excel chart insertion to summarize schedules and text grids before charting', () => {
    const chart = registryEntryForKindAndSurface('insert-chart', 'excel');
    if (!chart) throw new Error('missing Excel insert-chart capability registry entry');
    expect(chart.sequence.join('\n')).toContain('chart-ready summary table');
    expect(chart.sequence.join('\n')).toContain('pie only for <=6 non-negative parts');
    expect(chart.examples).toEqual(
      expect.arrayContaining([
        'chart bar \'Daily schedule\'!K6:L18 title="Weekly Hours by Activity" series=columns',
      ]),
    );
    expect(chart.failureModes).toContain('source range needs a derived summary table');
  });

  it('covers every Microsoft 365 surface with progressive-disclosure metadata', () => {
    expect(registryEntriesForSurface('word').map((entry) => entry.kind)).toEqual(
      expect.arrayContaining([
        'apply-style',
        'insert-table',
        'insert-content-control',
        'insert-hyperlink',
        'find-replace',
      ]),
    );
    expect(registryEntriesForSurface('powerpoint').map((entry) => entry.kind)).toEqual(
      expect.arrayContaining([
        'set-shape-text',
        'insert-image',
        'format-shape',
        'add-shape',
        'add-table-slide',
        'apply-slide-layout',
      ]),
    );
    expect(registryEntriesForSurface('outlook').map((entry) => entry.kind)).toEqual(
      expect.arrayContaining([
        'reply-mail',
        'create-mail',
        'add-attachment',
        'set-body',
        'set-recipients',
        'set-subject',
        'add-categories',
        'compose-appointment',
      ]),
    );
    expect(registryEntriesForSurface('onenote').map((entry) => entry.kind)).toEqual(
      expect.arrayContaining([
        'append-page',
        'set-page-title',
        'add-outline',
        'append-rich-text',
        'add-note-tag',
        'create-section',
      ]),
    );
    expect(registryEntriesForSurface('teams').map((entry) => entry.kind)).toEqual(
      expect.arrayContaining([
        'post-message',
        'post-card',
        'post-channel-message',
        'update-message',
        'create-online-meeting',
      ]),
    );
  });

  it('has no duplicate surface/kind entries', () => {
    const keys = capabilityRegistryEntries().map((entry) => `${entry.surface}:${entry.kind}`);
    expect(new Set(keys).size).toBe(keys.length);
  });
});
