import { afterEach, describe, expect, it, vi } from 'vitest';
import { getOfficeLoginHint } from './office-login-hint.js';

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('getOfficeLoginHint', () => {
  it('returns the UPN when the host answers', async () => {
    vi.stubGlobal('Office', {
      auth: { getAuthContext: () => Promise.resolve({ userPrincipalName: ' a@b.c ' }) },
    });
    expect(await getOfficeLoginHint()).toBe('a@b.c');
  });

  it('gives up after the timeout when the host never answers (PowerPoint for the web)', async () => {
    vi.stubGlobal('Office', { auth: { getAuthContext: () => new Promise(() => {}) } });
    const started = Date.now();
    expect(await getOfficeLoginHint(50)).toBeUndefined();
    expect(Date.now() - started).toBeLessThan(1000);
  });

  it('returns undefined when the API is missing or rejects', async () => {
    vi.stubGlobal('Office', {});
    expect(await getOfficeLoginHint()).toBeUndefined();
    vi.stubGlobal('Office', { auth: { getAuthContext: () => Promise.reject(new Error('no')) } });
    expect(await getOfficeLoginHint()).toBeUndefined();
  });
});
