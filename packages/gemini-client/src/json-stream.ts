/**
 * Discovery Engine's streaming REST methods return a JSON **array** whose elements
 * arrive incrementally: `[ {chunk}, {chunk}, ... ]`. This parses a byte stream into
 * the top-level objects as soon as each one completes, tolerating the array framing,
 * whitespace, and commas, and respecting strings/escapes so braces inside strings
 * don't confuse the depth counter.
 *
 * Scan state (`pos`, `depth`, string flags) persists across chunks; the buffer is
 * sliced only when a complete object is emitted, so characters are never re-scanned.
 */
/**
 * Framing facts the caller may need after the generator finishes. `closed` is true only when the
 * top-level array's closing `]` was seen outside any string or object — the one framing-level
 * proof that the server finished the response rather than the connection simply ending.
 * `incomplete` is true when non-whitespace bytes were left unparsed (a truncated object).
 */
export interface JsonStreamStatus {
  closed: boolean;
  incomplete: boolean;
}

export async function* parseJsonArrayStream(
  stream: ReadableStream<Uint8Array>,
  status?: JsonStreamStatus,
): AsyncGenerator<unknown> {
  const reader = stream.getReader();
  const decoder = new TextDecoder();
  let buf = '';
  let pos = 0; // next index to scan in buf (persists across chunks)
  let depth = 0;
  let start = -1; // index where the current top-level object began
  let inString = false;
  let escaped = false;

  function* drain(): Generator<unknown> {
    while (pos < buf.length) {
      const ch = buf[pos]!;
      if (inString) {
        if (escaped) escaped = false;
        else if (ch === '\\') escaped = true;
        else if (ch === '"') inString = false;
      } else if (ch === '"') {
        inString = true;
      } else if (ch === ']' && depth === 0) {
        if (status) status.closed = true;
      } else if (ch === '{') {
        if (depth === 0) start = pos;
        depth++;
      } else if (ch === '}') {
        depth--;
        if (depth === 0 && start >= 0) {
          const objText = buf.slice(start, pos + 1);
          yield JSON.parse(objText);
          // Drop the consumed prefix so the buffer can't grow without bound.
          buf = buf.slice(pos + 1);
          pos = 0;
          start = -1;
          continue;
        }
      }
      pos++;
    }
  }

  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    yield* drain();
  }
  buf += decoder.decode();
  yield* drain();
  // An object that opened but never closed is a truncation; stray framing (`]`, commas) is not.
  if (status && (depth > 0 || inString || start >= 0)) status.incomplete = true;
}
