/**
 * Dependency-free token *estimate* for budgeting (context tray) and chunk sizing.
 * Not a real tokenizer — we can't ship tiktoken into an Office webview cheaply — but a
 * stable approximation: ~4 chars/token for prose, floored at the word count so short,
 * punctuation-dense text isn't under-counted. Always treat as an estimate, never a limit.
 */
export function estimateTokens(text: string): number {
  if (!text) return 0;
  const chars = text.length;
  const words = text.trim().length === 0 ? 0 : text.trim().split(/\s+/).length;
  return Math.max(words, Math.ceil(chars / 4));
}
