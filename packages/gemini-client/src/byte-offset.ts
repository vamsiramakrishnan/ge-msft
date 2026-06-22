/**
 * Discovery Engine grounding/citation spans report `startIndex`/`endIndex` as
 * UTF-8 **byte** offsets into the answer text (multi-byte characters — CJK, emoji
 * — therefore advance the index by more than one). Our accumulated answer is a JS
 * UTF-16 string, so a span must be converted to a UTF-16 character index before it
 * can be used with `answerText.slice(start, end)`.
 *
 * This builds a byte-length prefix map over the string's code points once, then
 * binary-searches it per offset. Each entry maps a cumulative UTF-8 byte count to
 * the UTF-16 char index at that boundary, so a byte offset that lands exactly on a
 * character boundary round-trips, and an offset in the middle of a multi-byte
 * sequence clamps to the surrounding character boundary.
 */
export class ByteOffsetMapper {
  /** Cumulative UTF-8 byte count at each character boundary (ascending). */
  private readonly byteAt: number[] = [0];
  /** UTF-16 char index at each character boundary, parallel to `byteAt`. */
  private readonly charAt: number[] = [0];

  constructor(text: string) {
    let bytes = 0;
    // Iterating the string yields whole code points (surrogate pairs stay intact),
    // so `cp.length` is the UTF-16 unit count (1 for BMP, 2 for astral/emoji).
    let charIndex = 0;
    for (const cp of text) {
      bytes += utf8Len(cp.codePointAt(0)!);
      charIndex += cp.length;
      this.byteAt.push(bytes);
      this.charAt.push(charIndex);
    }
  }

  /** Total UTF-8 byte length of the mapped string. */
  get byteLength(): number {
    return this.byteAt[this.byteAt.length - 1]!;
  }

  /**
   * Convert a UTF-8 byte offset to a UTF-16 character index. Offsets past the end
   * clamp to the string length; offsets that land mid-character clamp down to the
   * preceding character boundary (the grounded span never splits a character).
   */
  toCharIndex(byteOffset: number): number {
    if (byteOffset <= 0) return 0;
    if (byteOffset >= this.byteLength) return this.charAt[this.charAt.length - 1]!;
    // Largest boundary whose byte count is <= byteOffset.
    let lo = 0;
    let hi = this.byteAt.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.byteAt[mid]! <= byteOffset) lo = mid;
      else hi = mid - 1;
    }
    return this.charAt[lo]!;
  }
}

/** UTF-8 encoded byte length of a single Unicode code point. */
function utf8Len(codePoint: number): number {
  if (codePoint <= 0x7f) return 1;
  if (codePoint <= 0x7ff) return 2;
  if (codePoint <= 0xffff) return 3;
  return 4;
}

/** Coerce a wire index (int64 sent as string, or a number) to a finite integer. */
export function parseByteIndex(value: string | number | undefined): number | undefined {
  if (value === undefined) return undefined;
  const n = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * Convert a single byte-offset span to a UTF-16 char span, or `undefined` if the
 * span is missing/degenerate. Convenience wrapper over {@link ByteOffsetMapper}.
 */
export function byteOffsetToCharIndex(
  mapper: ByteOffsetMapper,
  startByte: string | number | undefined,
  endByte: string | number | undefined,
): { start: number; end: number } | undefined {
  const sb = parseByteIndex(startByte);
  const eb = parseByteIndex(endByte);
  if (sb === undefined || eb === undefined) return undefined;
  const start = mapper.toCharIndex(sb);
  const end = mapper.toCharIndex(eb);
  if (end <= start) return undefined;
  return { start, end };
}
