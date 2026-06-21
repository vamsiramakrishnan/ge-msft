import { describe, it, expect } from 'vitest';
import { discoveryEngineHost, proxyBase, searchUrl, type GeminiClientConfig } from './config.js';

describe('discoveryEngineHost — residency pin (no silent global fallback)', () => {
  it('builds the regional host', () => {
    expect(discoveryEngineHost('eu')).toBe('https://discoveryengine.eu.rep.googleapis.com');
  });
  it('only the explicit "global" selects the global host', () => {
    expect(discoveryEngineHost('global')).toBe('https://discoveryengine.googleapis.com');
  });
  it('throws on an empty location rather than silently going global', () => {
    expect(() => discoveryEngineHost('')).toThrow(/residency/i);
  });
});

describe('proxyBase — the federated token sink must be https', () => {
  it('accepts https and trims a trailing slash', () => {
    expect(proxyBase('https://proxy.acme.com/')).toBe('https://proxy.acme.com');
  });
  it('allows http only for localhost dev', () => {
    expect(proxyBase('http://localhost:8080')).toBe('http://localhost:8080');
  });
  it('rejects non-https remote proxies (token-leak guard)', () => {
    expect(() => proxyBase('http://evil.example.com')).toThrow(/https/i);
  });
  it('rejects a malformed url', () => {
    expect(() => proxyBase('not a url')).toThrow(/invalid proxyurl/i);
  });
});

describe('url builders route through the validated proxy', () => {
  const cfg: GeminiClientConfig = {
    assistant: { project: 'p', location: 'eu', engine: 'e' },
    proxyUrl: 'https://proxy.acme.com',
  };
  it('search uses the proxy base', () => {
    expect(searchUrl(cfg)).toBe('https://proxy.acme.com/search');
  });
});
