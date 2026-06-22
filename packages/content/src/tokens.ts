/**
 * Dependency-free token *estimate* for budgeting (context tray) and chunk sizing.
 *
 * Not a real tokenizer — we can't ship tiktoken into an Office webview cheaply — but a
 * stable, deterministic approximation. The heuristic is script-aware because byte/char
 * density per token varies wildly by script:
 *
 *   • Latin/whitespace-delimited prose: subword tokenizers (BPE/SentencePiece) average
 *     ~4 characters per token. We also floor at the whitespace word count so short,
 *     punctuation-dense text ("a. b! c?") isn't under-counted.
 *   • CJK (Han/Hiragana/Katakana/Hangul) and other space-free ideographic scripts:
 *     tokenizers emit roughly one token per character (often slightly more once you
 *     account for byte-fallback on rare glyphs). We count these at ~1 token/char.
 *
 * We classify per code point and sum the two populations, so a mixed-script string is
 * estimated correctly rather than averaged. Always treat the result as an estimate, never
 * a hard limit.
 */

/** Chars-per-token for whitespace-delimited (Latin-ish) text. */
const LATIN_CHARS_PER_TOKEN = 4;
/** Tokens-per-char for ideographic / space-free scripts. */
const CJK_TOKENS_PER_CHAR = 1;

/** Code-point ranges for scripts that tokenize at roughly one token per character. */
function isCjkCodePoint(cp: number): boolean {
  return (
    (cp >= 0x3000 && cp <= 0x303f) || // CJK symbols & punctuation
    (cp >= 0x3040 && cp <= 0x30ff) || // Hiragana + Katakana
    (cp >= 0x3400 && cp <= 0x4dbf) || // CJK Ext. A
    (cp >= 0x4e00 && cp <= 0x9fff) || // CJK Unified Ideographs
    (cp >= 0xf900 && cp <= 0xfaff) || // CJK Compatibility Ideographs
    (cp >= 0xac00 && cp <= 0xd7af) || // Hangul Syllables
    (cp >= 0xff00 && cp <= 0xffef) || // Halfwidth/Fullwidth forms
    (cp >= 0x20000 && cp <= 0x2ebef) // CJK Ext. B–F (supplementary plane)
  );
}

export function estimateTokens(text: string): number {
  if (!text) return 0;

  let cjkChars = 0;
  let otherChars = 0;
  // Iterate by code point so supplementary-plane CJK (Ext. B+) counts as one char.
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (isCjkCodePoint(cp)) cjkChars++;
    else otherChars++;
  }

  const cjkTokens = cjkChars * CJK_TOKENS_PER_CHAR;
  const latinTokens = Math.ceil(otherChars / LATIN_CHARS_PER_TOKEN);

  const trimmed = text.trim();
  const words = trimmed.length === 0 ? 0 : trimmed.split(/\s+/).length;

  return Math.max(words, cjkTokens + latinTokens);
}
