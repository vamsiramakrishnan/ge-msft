import { describe, it, expect } from 'vitest';
import type { Surface } from '@ge/contracts';
import { selectBridge, isSupportedSurface, SUPPORTED_SURFACES } from './select-bridge.js';

describe('selectBridge (the surface seam)', () => {
  it('constructs the matching bridge for each wired surface', () => {
    expect(selectBridge('word')?.surface).toBe('word');
    expect(selectBridge('excel')?.surface).toBe('excel');
    expect(selectBridge('outlook')?.surface).toBe('outlook');
    expect(selectBridge('powerpoint')?.surface).toBe('powerpoint');
    expect(selectBridge('onenote')?.surface).toBe('onenote');
    expect(selectBridge('teams')?.surface).toBe('teams');
  });

  it('returns undefined for an unknown / undetected surface', () => {
    expect(selectBridge(undefined)).toBeUndefined();
    expect(selectBridge('mystery' as Surface)).toBeUndefined();
  });

  it('passes Teams options through to the bridge', () => {
    const bridge = selectBridge('teams', {
      teams: { transcript: { transcript: 'a: hello', meetingTitle: 'Sync' } },
    });
    expect(bridge?.surface).toBe('teams');
  });

  it('isSupportedSurface agrees with selectBridge', () => {
    const all: Surface[] = ['word', 'excel', 'outlook', 'teams', 'powerpoint', 'onenote'];
    for (const s of all) {
      expect(isSupportedSurface(s)).toBe(selectBridge(s) !== undefined);
    }
    expect(isSupportedSurface(undefined)).toBe(false);
  });

  it('SUPPORTED_SURFACES lists exactly the wired surfaces', () => {
    expect([...SUPPORTED_SURFACES].sort()).toEqual([
      'excel',
      'onenote',
      'outlook',
      'powerpoint',
      'teams',
      'word',
    ]);
  });
});
