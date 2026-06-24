import { describe, it, expect, afterEach } from 'vitest';
import { contentHash } from './hash.js';

/**
 * The fallback FNV-1a path only runs when Web Crypto's subtle digest is unavailable.
 * We temporarily remove crypto.subtle to exercise it, then restore it.
 */
const realCrypto = globalThis.crypto;

afterEach(() => {
  Object.defineProperty(globalThis, 'crypto', { value: realCrypto, configurable: true });
});

function removeSubtle(): void {
  // Replace the crypto object with one that has no subtle, forcing the fallback.
  Object.defineProperty(globalThis, 'crypto', { value: {}, configurable: true });
}

describe('contentHash — Web Crypto path', () => {
  it('matches the known SHA-256 vector for "abc"', async () => {
    expect(await contentHash('abc')).toBe(
      'sha256:ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
    );
  });

  it('hashes empty string to the canonical SHA-256 empty digest', async () => {
    expect(await contentHash('')).toBe(
      'sha256:e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
    );
  });

  it('is deterministic and content-sensitive (avalanche on a single char)', async () => {
    expect(await contentHash('hello world')).toBe(await contentHash('hello world'));
    expect(await contentHash('hello world')).not.toBe(await contentHash('hello worle'));
  });

  it('encodes UTF-8 bytes (multibyte content produces a distinct, valid digest)', async () => {
    const h = await contentHash('café ☕ 日本');
    expect(h).toMatch(/^sha256:[0-9a-f]{64}$/);
    // Differs from the ASCII-only prefix of the same logical text.
    expect(h).not.toBe(await contentHash('cafe coffee'));
  });
});

describe('contentHash — degraded FNV-1a fallback (no crypto.subtle)', () => {
  it('falls back to an fnv1a32-prefixed digest auditable as non-crypto', async () => {
    removeSubtle();
    const h = await contentHash('abc');
    expect(h).toMatch(/^fnv1a32:[0-9a-f]{8}$/);
    // Must NOT masquerade as a SHA-256 record.
    expect(h.startsWith('sha256:')).toBe(false);
  });

  it('matches the canonical FNV-1a 32-bit vector for "abc"', async () => {
    // FNV-1a/32 of "abc" is 0x1a47e90b (well-known reference value).
    removeSubtle();
    expect(await contentHash('abc')).toBe('fnv1a32:1a47e90b');
  });

  it('is deterministic and content-sensitive in the fallback path', async () => {
    removeSubtle();
    expect(await contentHash('provenance')).toBe(await contentHash('provenance'));
    expect(await contentHash('provenance')).not.toBe(await contentHash('Provenance'));
  });

  it('hashes the empty string to the FNV-1a offset basis', async () => {
    // The 32-bit FNV offset basis is 0x811c9dc5; an empty input returns it unchanged.
    removeSubtle();
    expect(await contentHash('')).toBe('fnv1a32:811c9dc5');
  });
});
