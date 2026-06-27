import { describe, it, expect } from 'vitest';
import { surfaceFromHost, detectSurface } from './host.js';

describe('surfaceFromHost', () => {
  it('maps known Office hosts to surfaces', () => {
    expect(surfaceFromHost('Word')).toBe('word');
    expect(surfaceFromHost('Excel')).toBe('excel');
    expect(surfaceFromHost('PowerPoint')).toBe('powerpoint');
    expect(surfaceFromHost('OneNote')).toBe('onenote');
    expect(surfaceFromHost('Outlook')).toBe('outlook');
  });

  it('returns undefined for unknown / missing hosts', () => {
    expect(surfaceFromHost('Visio')).toBeUndefined();
    expect(surfaceFromHost(undefined)).toBeUndefined();
    expect(surfaceFromHost(null)).toBeUndefined();
  });
});

describe('detectSurface', () => {
  it('detects Outlook by the mailbox object before host', () => {
    expect(detectSurface({ context: { mailbox: {} } })).toBe('outlook');
  });
  it('detects editors by host', () => {
    expect(detectSurface({ context: { host: 'Excel' } })).toBe('excel');
  });
  it('returns undefined without an Office context', () => {
    expect(detectSurface({})).toBeUndefined();
  });
});
