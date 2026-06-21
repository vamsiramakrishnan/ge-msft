import { describe, it, expect, vi } from 'vitest';
import {
  NaaAuthClient,
  type MsalLike,
  type MsalAccount,
  type MsalAuthResult,
} from './auth-client.js';

const account: MsalAccount = { username: 'v.k@acme', name: 'V K', localAccountId: 'oid-1' };

function fakeMsal(over: Partial<MsalLike> = {}): MsalLike {
  return {
    getActiveAccount: () => account,
    acquireTokenSilent: (req) =>
      Promise.resolve<MsalAuthResult>({
        accessToken: `acc:${req.scopes.join(',')}`,
        idToken: 'id-tok',
        account,
      }),
    ...over,
  };
}

const opts = { idTokenScopes: ['api://ge/.default'], graphScopes: ['Mail.Read'] };

describe('NaaAuthClient', () => {
  it('returns the id token (WIF subject) from silent acquisition', async () => {
    const auth = new NaaAuthClient(fakeMsal(), opts);
    expect(await auth.getIdToken()).toBe('id-tok');
  });

  it('throws when no id token comes back', async () => {
    const auth = new NaaAuthClient(
      fakeMsal({ acquireTokenSilent: () => Promise.resolve({ accessToken: 'a' }) }),
      opts,
    );
    await expect(auth.getIdToken()).rejects.toThrow(/no id token/i);
  });

  it('acquires a Graph token for the requested scopes', async () => {
    const auth = new NaaAuthClient(fakeMsal(), opts);
    expect(await auth.getGraphToken(['Calendars.Read'])).toBe('acc:Calendars.Read');
  });

  it('falls back to the configured graph scopes when none are passed', async () => {
    const auth = new NaaAuthClient(fakeMsal(), opts);
    expect(await auth.getGraphToken([])).toBe('acc:Mail.Read');
  });

  it('falls back to popup when silent acquisition needs interaction', async () => {
    const interaction = Object.assign(new Error('need ui'), {
      name: 'InteractionRequiredAuthError',
    });
    const popup = vi.fn(() =>
      Promise.resolve<MsalAuthResult>({ accessToken: 'popup-acc', idToken: 'popup-id' }),
    );
    const auth = new NaaAuthClient(
      fakeMsal({ acquireTokenSilent: () => Promise.reject(interaction), acquireTokenPopup: popup }),
      opts,
    );
    expect(await auth.getIdToken()).toBe('popup-id');
    expect(popup).toHaveBeenCalledOnce();
  });

  it('does not pop up for non-interaction errors', async () => {
    const popup = vi.fn();
    const auth = new NaaAuthClient(
      fakeMsal({
        acquireTokenSilent: () => Promise.reject(new Error('network')),
        acquireTokenPopup: popup,
      }),
      opts,
    );
    await expect(auth.getIdToken()).rejects.toThrow('network');
    expect(popup).not.toHaveBeenCalled();
  });

  it('derives identity from the active account', async () => {
    const auth = new NaaAuthClient(fakeMsal(), opts);
    expect(await auth.getIdentity()).toEqual({
      username: 'v.k@acme',
      displayName: 'V K',
      oid: 'oid-1',
    });
  });
});
