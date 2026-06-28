import type {
  ActuationKind,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import { buildDocStateSnapshot } from '@ge/content';
import { ONENOTE_CAPABILITIES } from './capabilities.js';
import { isSet } from './capabilities-runtime.js';
import {
  pageElementToDocStateBlocks,
  pageToContext,
  searchPage,
  type PageElement,
} from './capture.js';
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
/**
 * The exact `ActuationKind`s {@link OneNoteBridge.actuate} handles (ADR-0006 closure source of
 * truth). The conformance test asserts this equals the advertised manifest's actuation kinds.
 */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = ['append-page'];

export class OneNoteBridge implements DocBridge {
  readonly surface = 'onenote' as const;

  /** Monotonic `<doc_state>` version, bumped on each capture (ADR-0003 Layer B element 1). */
  private docStateVersion = 0;

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
    const page = await this.readActivePage();
    return page ? pageToContext(page) : [];
  }

  canRevealContext(ref: ContextRef): boolean {
    return (
      ref.surface === 'onenote' && ref.kind === 'page' && onenoteRevealTarget(ref) !== undefined
    );
  }

  async revealContext(ref: ContextRef): Promise<void> {
    const target = onenoteRevealTarget(ref);
    if (!target || !isSet('OneNoteApi', '1.1')) return;
    await OneNote.run(async (ctx) => {
      if (target.clientUrl) {
        ctx.application.navigateToPageWithClientUrl(target.clientUrl);
        await ctx.sync();
        return;
      }
      const page = ctx.application.getActivePageOrNull();
      page.load('id');
      await ctx.sync();
      if (page.isNullObject || (target.pageId && page.id !== target.pageId)) return;
      ctx.application.navigateToPage(page);
      await ctx.sync();
    });
  }

  /**
   * ADR-0003 Layer B element 1 / ADR-0006 `outline` read: an ambient structural snapshot of the
   * active page — its title (heading) + paragraph outline, mapped through the same native blocks
   * grounding context uses. OneNote is web-only with a narrow API (no whole-notebook enumeration),
   * so the "document" is the active page. Older host / no active page / empty page → `undefined`
   * (the runtime just streams without the ambient part). Version increments per capture.
   */
  async captureDocState(): Promise<DocStateSnapshot | undefined> {
    const page = await this.readActivePage();
    if (!page) return undefined;
    const blocks = pageElementToDocStateBlocks(page);
    if (blocks.length === 0) return undefined;
    this.docStateVersion += 1;
    return buildDocStateSnapshot({
      surface: 'onenote',
      version: this.docStateVersion,
      ...(page.title.trim() ? { title: page.title } : {}),
      blocks,
    });
  }

  /**
   * ADR-0006 `search` read: scan the active page's paragraphs for `query` and return the matching
   * paragraphs as `ResolvedContext` data (never instructions), bounded by `searchPage`. Older host /
   * empty query / no active page / no match → `[]`.
   */
  async searchDocument(query: string): Promise<ResolvedContext[]> {
    const q = query.trim();
    if (!q) return [];
    const page = await this.readActivePage();
    return page ? searchPage(page, q) : [];
  }

  /**
   * Read the active OneNote page into a pure {@link PageElement} — the shared host read behind
   * `resolveContext`/`captureDocState`/`searchDocument`. Read-only: loads the page title + its
   * outlines' rich-text paragraphs in the fewest syncs (per the "batch loads before sync" rule) and
   * writes nothing. Older host / no active page → `undefined`.
   */
  private async readActivePage(): Promise<PageElement | undefined> {
    if (!isSet('OneNoteApi', '1.1')) return undefined;
    return OneNote.run(async (ctx) => {
      const page = ctx.application.getActivePageOrNull();
      page.load('id,title');
      await ctx.sync();
      if (page.isNullObject) return undefined;

      // Outlines → paragraphs → rich text, batched into the fewest syncs: (1) load the page's
      // contents/outline types, (2) queue every outline's paragraphs in one sync, then (3) queue
      // every rich-text body in one sync.
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

      return {
        pageId: page.id,
        title: page.title,
        paragraphs: texts.map((t) => t.text ?? '').filter((t) => t.trim().length > 0),
      };
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

interface OneNoteRevealTarget {
  pageId?: string;
  clientUrl?: string;
}

function onenoteRevealTarget(ref: ContextRef): OneNoteRevealTarget | undefined {
  if (ref.surface !== 'onenote' || ref.kind !== 'page') return undefined;
  const url = prefixedValue(ref.anchor?.locator, 'clientUrl:', 'url:');
  if (url) return { clientUrl: url };
  const pageId =
    prefixedValue(ref.id, 'on:page:', 'onenote:page:') ??
    prefixedValue(ref.anchor?.locator, 'page:', 'on:page:', 'onenote:page:');
  return pageId ? { pageId } : undefined;
}

function prefixedValue(value: string | undefined, ...prefixes: string[]): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) return undefined;
  for (const prefix of prefixes) {
    if (trimmed.startsWith(prefix)) {
      const rest = trimmed.slice(prefix.length).trim();
      return rest || undefined;
    }
  }
  return undefined;
}
