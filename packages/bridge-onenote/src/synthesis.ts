import type { SourceRef } from '@ge/contracts';

/**
 * Pure builders for OneNote page synthesis — no Office.js, so unit-testable. OneNote's
 * `Page.addOutline` / `Section.addPage` accept an HTML string (the supported-HTML subset); the
 * synthesis is rendered as paragraphs, each carrying an **inline citation tag** per claim so the
 * page stays traceable to its grounding sources (the mockup's per-claim `[source]` chips). All
 * untrusted text is HTML-escaped before it lands on the page.
 */

/** One synthesized claim and the source it grounds on. */
export interface SynthesisPart {
  heading?: string;
  text: string;
  source?: SourceRef;
}

/** Escape the five HTML-significant characters so host/agent text is rendered as data. */
export function escapeHtml(input: string): string {
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Render a single source as an inline citation tag (title + optional locator). */
export function citationTag(source: SourceRef): string {
  const label = source.locator ? `${source.title} · ${source.locator}` : source.title;
  return `<span data-ge-cite="1">[${escapeHtml(label)}]</span>`;
}

/** Render one synthesis part as a heading (optional) + a paragraph with a trailing citation. */
export function partToHtml(part: SynthesisPart): string {
  const head = part.heading && part.heading.trim() ? `<h2>${escapeHtml(part.heading)}</h2>` : '';
  const cite = part.source ? ` ${citationTag(part.source)}` : '';
  return `${head}<p>${escapeHtml(part.text)}${cite}</p>`;
}

/** Build the full page-body HTML (an outline) from the synthesized parts. */
export function buildPageHtml(parts: SynthesisPart[]): string {
  return parts.map(partToHtml).join('');
}
