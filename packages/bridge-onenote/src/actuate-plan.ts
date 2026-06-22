import type { ActuationRequest } from '@ge/contracts';
import { buildPageHtml, escapeHtml, type SynthesisPart } from './synthesis.js';

/**
 * Pure translation of an `append-page` actuation into a host plan — testable without Office.js.
 *
 * The agent supplies the synthesis as either:
 *   - `params.html`: a prebuilt body (already citation-tagged); used verbatim, OR
 *   - `params.text` + `params.sources`: plain text we render to a single citation-tagged
 *     paragraph (one tag per supplied source).
 *
 * The page title is taken from the target's `matchText` (a caller-chosen title) or defaults to a
 * timestamp-free "Synthesis" label; the bridge passes the title to `Section.addPage` and the
 * body HTML to `Page.addOutline`.
 */

export interface AppendPagePlan {
  title: string;
  html: string;
}

export function planAppendPage(req: ActuationRequest): AppendPagePlan {
  const p = req.params;
  const title = p.target?.matchText?.trim() || 'Synthesis';

  if (p.html && p.html.trim()) {
    return { title, html: p.html };
  }

  const text = p.text ?? '';
  const sources = p.sources ?? [];
  const parts: SynthesisPart[] = text.trim()
    ? [{ text, ...(sources[0] ? { source: sources[0] } : {}) }]
    : [];
  // If there are multiple sources, append the remaining citation tags after the paragraph.
  let html = buildPageHtml(parts);
  if (sources.length > 1) {
    const extra = sources
      .slice(1)
      .map(
        (s) =>
          `<span data-ge-cite="1">[${escapeHtml(s.locator ? `${s.title} · ${s.locator}` : s.title)}]</span>`,
      )
      .join(' ');
    html += `<p>${extra}</p>`;
  }
  return { title, html };
}
