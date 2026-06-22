import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import { ONENOTE_CAPABILITIES } from './capabilities.js';
import { isSet } from './capabilities-runtime.js';
import { pageToContext, type PageElement } from './capture.js';
import { planAppendPage } from './actuate-plan.js';

/**
 * The OneNote `DocBridge`. The ONLY place Office.js (`OneNote.run`) is touched. OneNote is
 * **web-only** and ships behind the **legacy XML add-in manifest** (`manifests/onenote.manifest.xml`),
 * with a narrower API than the unified-manifest hosts — so this bridge scopes to what the
 * `OneNoteApi` surface actually exposes: reading the active page's title + outline rich text, and
 * writing a new synthesized page (`append-page`) whose claims carry inline citation tags. Pure
 * mapping lives in `capture.ts` / `synthesis.ts` / `actuate-plan.ts` (unit-tested); this file is
 * the host wiring. There is no OneNote object-model event API in this set, so `watch` is omitted.
 *
 * Requirement set (confirmed against node_modules/@types/office-js/index.d.ts):
 *   - `OneNote.Application.getActivePage()` / `getActiveSection()` → OneNoteApi 1.1 (l.168508 / l.168536).
 *   - `Page.title` / `Page.contents` (PageContentCollection) → 1.1 (l.170264 / l.170215);
 *     `PageContent.outline` → 1.1 (l.170498); `Outline.paragraphs` → 1.1 (l.170674);
 *     `Paragraph.richText` → 1.1 (l.170835); `RichText.text` → 1.1 (l.171198).
 *   - `Section.addPage(title)` → 1.1 (l.170042); `Page.addOutline(left, top, html)` → 1.1 (l.170290).
 */
export class OneNoteBridge implements DocBridge {
  readonly surface = 'onenote' as const;

  getCapabilities(): CapabilityManifest {
    return ONENOTE_CAPABILITIES;
  }

  async listContext(): Promise<ContextRef[]> {
    if (!isSet('OneNoteApi', '1.1')) return [];
    return OneNote.run(async (ctx) => {
      const page = ctx.application.getActivePageOrNull();
      page.load('id,title');
      await ctx.sync();
      if (page.isNullObject) return [];
      return [
        {
          id: `on:page:${page.id}`,
          kind: 'page' as const,
          surface: 'onenote' as const,
          title: page.title || 'Current page',
          live: true,
        },
      ];
    });
  }

  async resolveContext(_ref: ContextRef): Promise<ResolvedContext[]> {
    if (!isSet('OneNoteApi', '1.1')) return [];
    return OneNote.run(async (ctx) => {
      const page = ctx.application.getActivePageOrNull();
      page.load('id,title');
      await ctx.sync();
      if (page.isNullObject) return [];

      // Outlines → paragraphs → rich text, batched into the fewest syncs (per the wave-1
      // "batch loads before sync" rule): (1) load the page's contents/outline types, (2) queue
      // every outline's paragraphs in one sync, then (3) queue every rich-text body in one sync.
      const contents = page.contents;
      contents.load('items/type');
      await ctx.sync();

      const paragraphCollections: OneNote.ParagraphCollection[] = [];
      for (const content of contents.items) {
        if (content.type !== 'Outline') continue;
        const paragraphs = content.outline.paragraphs;
        paragraphs.load('items/type');
        paragraphCollections.push(paragraphs);
      }
      await ctx.sync();

      const texts: OneNote.RichText[] = [];
      for (const paragraphs of paragraphCollections) {
        for (const para of paragraphs.items) {
          if (para.type !== 'RichText') continue;
          const rt = para.richText;
          rt.load('text');
          texts.push(rt);
        }
      }
      await ctx.sync();

      const element: PageElement = {
        pageId: page.id,
        title: page.title,
        paragraphs: texts.map((t) => t.text ?? '').filter((t) => t.trim().length > 0),
      };
      return pageToContext(element);
    });
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    if (req.kind !== 'append-page') {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'unsupported', message: `OneNote bridge cannot ${req.kind}` },
      };
    }
    const plan = planAppendPage(req);
    if (!plan.html.trim()) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'empty_synthesis', message: 'append-page needs params.html or params.text' },
      };
    }
    if (!isSet('OneNoteApi', '1.1')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'onenote_unsupported',
          message: 'OneNote write API is unavailable on this host.',
        },
      };
    }
    return OneNote.run(async (ctx) => {
      const section = ctx.application.getActiveSectionOrNull();
      section.load('id');
      await ctx.sync();
      if (section.isNullObject) {
        return {
          ok: false,
          changeId: req.changeId,
          kind: req.kind,
          degraded: true,
          error: { code: 'no_section', message: 'No active OneNote section to add the page to.' },
        };
      }
      const page = section.addPage(plan.title);
      // The outline is positioned at a standard top-left offset (points) on the new page.
      page.addOutline(40, 40, plan.html);
      page.load('id');
      await ctx.sync();
      return { ok: true, changeId: req.changeId, kind: req.kind, location: `page:${page.id}` };
    });
  }
}
