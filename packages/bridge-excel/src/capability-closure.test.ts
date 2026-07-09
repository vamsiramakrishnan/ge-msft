import { describe, expect, it } from 'vitest';
import { implementedRegistryKindsForSurface } from '../../contracts/src/capability-registry.js';
import { EXCEL_CAPABILITIES } from './capabilities.js';
import {
  ExcelBridge,
  HANDLED_ACTUATIONS,
  MAX_READ_CELLS,
  isA1Address,
  withinReadBudget,
} from './excel-bridge.js';

/**
 * ADR-0006 capability closure — Excel. Self-contained (no runtime/contracts closure-helper import):
 * the advertised manifest and what `actuate()` handles must be the SAME set, and the advertised
 * `reads` must match the implemented read ports — including the newly-added addressable `readRange`.
 */
describe('Excel capability closure', () => {
  it('advertised actuation kinds === handled actuation kinds', () => {
    const advertised = new Set(EXCEL_CAPABILITIES.actuations.map((a) => a.kind));
    const handled = new Set(HANDLED_ACTUATIONS);
    expect(advertised).toEqual(handled);
  });

  it('advertised actuation kinds === registry implemented Excel capabilities', () => {
    const advertised = new Set(EXCEL_CAPABILITIES.actuations.map((a) => a.kind));
    expect(advertised).toEqual(new Set(implementedRegistryKindsForSurface('excel')));
  });

  it('advertised reads match the implemented read ports', () => {
    expect(new Set(EXCEL_CAPABILITIES.reads)).toEqual(new Set(['outline', 'read', 'search']));
  });

  it('exposes a read port for each advertised read verb (outline/read/search)', () => {
    const bridge = new ExcelBridge();
    expect(typeof bridge.captureDocState).toBe('function'); // outline
    expect(typeof bridge.readRange).toBe('function'); // addressable read <A1|NamedRange>
    expect(typeof bridge.searchDocument).toBe('function'); // search
  });
});

describe('Excel readRange selector classification (pure)', () => {
  it('treats A1 / sheet-qualified / absolute references as addresses', () => {
    expect(isA1Address('A1')).toBe(true);
    expect(isA1Address('A1:B3')).toBe(true);
    expect(isA1Address('$A$1')).toBe(true);
    expect(isA1Address('Sheet1!A1:B3')).toBe(true);
    expect(isA1Address("'My Sheet'!C2")).toBe(true);
  });

  it('treats non-A1 selectors as named ranges', () => {
    expect(isA1Address('Revenue')).toBe(false);
    expect(isA1Address('Q1_Totals')).toBe(false);
    expect(isA1Address('')).toBe(false);
  });
});

describe('Excel read budget (pure) — bound checked before .values is materialized', () => {
  it('accepts ranges within MAX_READ_CELLS and rejects oversized / empty ones', () => {
    expect(withinReadBudget(10, 5)).toBe(true); // 50 cells
    expect(withinReadBudget(100, 100)).toBe(true); // 10_000 == cap
    expect(withinReadBudget(1_048_576, 1)).toBe(false); // a whole-column-sized finite range
    expect(withinReadBudget(MAX_READ_CELLS + 1, 1)).toBe(false);
    expect(withinReadBudget(0, 0)).toBe(false); // empty / unresolved
    expect(withinReadBudget(undefined, undefined)).toBe(false);
  });
});
