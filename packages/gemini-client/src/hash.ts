/**
 * Cryptographic content hash for provenance stamping. The provenance record carries a
 * hash of the generated content so a reviewer can detect post-hoc edits; using SHA-256
 * (rather than a cheap non-crypto digest) makes that tamper-evidence collision-resistant.
 *
 * Uses Web Crypto (`crypto.subtle.digest`) — available in every add-in webview and in
 * Node 20+ via `globalThis.crypto`. It is therefore **async**. We deliberately avoid
 * Node's `crypto` module so the client stays browser-safe.
 */

const SHA256_PREFIX = 'sha256:';

/**
 * Return `sha256:<hex>` for the UTF-8 bytes of `text`. Async because Web Crypto's
 * digest is a Promise. Falls back to a clearly-labelled non-crypto digest only if
 * `crypto.subtle` is genuinely unavailable (it should not be in supported runtimes).
 */
export async function contentHash(text: string): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) return fallbackHash(text);
  const bytes = new TextEncoder().encode(text);
  const digest = await subtle.digest('SHA-256', bytes);
  return SHA256_PREFIX + toHex(new Uint8Array(digest));
}

function toHex(bytes: Uint8Array): string {
  let out = '';
  for (const b of bytes) out += b.toString(16).padStart(2, '0');
  return out;
}

/**
 * Degraded digest used only when Web Crypto is missing. FNV-1a, 32-bit — NOT
 * collision-resistant; the distinct `fnv1a32:` prefix makes such records auditable as
 * having been produced without a cryptographic hash.
 */
function fallbackHash(text: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return 'fnv1a32:' + (h >>> 0).toString(16).padStart(8, '0');
}
