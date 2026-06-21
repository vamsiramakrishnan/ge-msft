/**
 * Small, dependency-free content hash for provenance stamping. Not a security
 * primitive — it identifies *which* generated content a provenance record covers,
 * so a reviewer can detect post-hoc edits. FNV-1a, 32-bit, hex.
 */
export function contentHash(text: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16).padStart(8, '0');
}
