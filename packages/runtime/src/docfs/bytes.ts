/**
 * Byte-accurate text budgeting shared by every DocFs mount. `ReadOpts.maxBytes` (@ge/contracts) is
 * documented as a byte budget; a naive `text.slice(0, maxBytes)` under-truncates non-ASCII text
 * (multi-byte UTF-8 characters make a slice of N JS characters exceed N bytes), which silently
 * breaks that contract for CJK/emoji/accented content. This truncates on real UTF-8 byte boundaries.
 */
export function byteLength(text: string): number {
  return new TextEncoder().encode(text).length;
}

/**
 * Truncate `text` to at most `maxBytes` UTF-8 bytes, landing on a real codepoint boundary.
 *
 * A naive `TextDecoder().decode(bytes.slice(0, maxBytes))` in its default (non-fatal) mode
 * SILENTLY REPLACES an incomplete trailing multi-byte sequence with one U+FFFD replacement
 * character — which is itself 3 UTF-8 bytes, so re-encoding the "truncated" result can exceed
 * `maxBytes` (e.g. cutting 1 byte into a 3-byte sequence still yields a 3-byte replacement char,
 * growing back past the budget). Using `fatal: true` instead makes an incomplete tail throw, so we
 * can back off byte-by-byte (at most 3 times — the longest UTF-8 sequence is 4 bytes) to the nearest
 * complete boundary, guaranteeing the result never exceeds `maxBytes`.
 */
export function truncateToBytes(
  text: string,
  maxBytes: number,
): { text: string; truncated: boolean } {
  const bytes = new TextEncoder().encode(text);
  if (bytes.length <= maxBytes) return { text, truncated: false };
  const decoder = new TextDecoder('utf-8', { fatal: true });
  for (let back = 0; back <= 3 && maxBytes - back >= 0; back++) {
    try {
      return { text: decoder.decode(bytes.slice(0, maxBytes - back)), truncated: true };
    } catch {
      // Incomplete trailing sequence — back off one more byte and retry.
    }
  }
  return { text: '', truncated: true };
}
