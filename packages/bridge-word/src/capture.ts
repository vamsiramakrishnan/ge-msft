import type { ContextRef, ResolvedContext } from '@ge/contracts';
import {
  native,
  toContextNative,
  type Block,
  type NativeContent,
  type ToContextOptions,
} from '@ge/content';
import type { WordComment, WordParagraph } from './host-port.js';

/** A re-resolved search hit at the host boundary: the matched text + a short surrounding hint. */
export interface WordSearchHit {
  readonly text: string;
  readonly contextHint?: string;
}

/**
 * Pure mapping from Word's native object model into grounding-ready context — no Office.js
 * here, so it's unit-testable. The `WordBridge` reads paragraphs/tables/content-controls via
 * `Word.run` and hands the extracted elements to these functions; they go straight through
 * `@ge/content` (native path, no Markdown round-trip) and carry a `cc:<id>` write-back locator.
 */
export interface WordElement {
  kind: 'heading' | 'paragraph' | 'table';
  text: string;
  level?: number; // heading level, derived from the built-in style
  contentControlId?: number;
  columns?: string[];
  rows?: string[][];
}

export function wordElementsToBlocks(elements: WordElement[]): Block[] {
  const blocks: Block[] = [];
  for (const el of elements) {
    const locator = el.contentControlId !== undefined ? `cc:${el.contentControlId}` : undefined;
    if (el.kind === 'heading') {
      blocks.push(native.heading(el.text, el.level ?? 2, locator));
    } else if (el.kind === 'table' && el.columns && el.rows) {
      blocks.push(native.table({ columns: el.columns, rows: el.rows }, locator));
    } else {
      blocks.push(native.paragraph(el.text, locator));
    }
  }
  return blocks;
}

export function wordDocumentToContext(
  sourceId: string,
  title: string | undefined,
  elements: WordElement[],
  opts: ToContextOptions = {},
): ResolvedContext[] {
  const content: NativeContent = {
    sourceId,
    surface: 'word',
    ...(title ? { title } : {}),
    blocks: wordElementsToBlocks(elements),
  };
  return toContextNative(content, opts);
}

/** A live selection is a single text part (re-resolved at send-time). */
export function wordSelectionToContext(text: string): ResolvedContext[] {
  if (!text.trim()) return [];
  return [
    {
      ref: {
        id: 'word:selection',
        kind: 'selection',
        surface: 'word',
        title: 'Selection',
        preview: text.slice(0, 120),
        live: true,
      },
      value: { as: 'text', text, mimeType: 'text/markdown' },
    },
  ];
}

/** Map a Word built-in style name to a heading level (0 = not a heading). */
export function headingLevel(styleBuiltIn: string): number {
  const m = /Heading\s*(\d)/i.exec(styleBuiltIn);
  return m ? Number(m[1]) : 0;
}

/**
 * Map read-back body paragraphs (text + built-in style) to {@link WordElement}s — headings carry
 * their derived level — exactly as `WordBridge.resolveContext` does. Shared so the `<doc_state>`
 * snapshot's blocks come from the same native mapping as grounding context (ADR-0003).
 */
export function paragraphsToElements(paras: readonly WordParagraph[]): WordElement[] {
  return paras.map((p) => {
    const level = headingLevel(p.styleBuiltIn);
    return level > 0
      ? { kind: 'heading' as const, text: p.text, level }
      : { kind: 'paragraph' as const, text: p.text };
  });
}

/** Body paragraphs → `Block[]` for `buildDocStateSnapshot` (headings get levels + locators). */
export function paragraphsToBlocks(paras: readonly WordParagraph[]): Block[] {
  return wordElementsToBlocks(paragraphsToElements(paras));
}

/**
 * Map bounded, re-resolved `body.search` hits to content-anchored {@link ResolvedContext} — one
 * live text part per hit, anchored by the matched text (the contextHint is folded into the part
 * text as a surrounding cue). Re-resolution happens at the host; this is the pure shaping step.
 * Empty input → `[]`.
 */
export function searchHitsToContext(
  query: string,
  hits: readonly WordSearchHit[],
): ResolvedContext[] {
  const out: ResolvedContext[] = [];
  for (const hit of hits) {
    const text = hit.text.trim();
    if (!text) continue;
    const body =
      hit.contextHint && hit.contextHint.trim() ? `${text}\n\n…${hit.contextHint.trim()}` : text;
    out.push({
      ref: {
        id: 'word:search',
        kind: 'selection',
        surface: 'word',
        title: `Match: ${query}`.slice(0, 80),
        preview: text.slice(0, 120),
        live: true,
        anchor: { matchText: text.slice(0, 120) },
      },
      value: { as: 'text', text: body, mimeType: 'text/markdown' },
    });
  }
  return out;
}

/** Longest anchored-text excerpt carried with a comment, so one huge range can't flood a turn. */
const MAX_COMMENT_ANCHOR_CHARS = 500;

function commentAuthor(name: string): string {
  return name.trim() || 'Unknown author';
}

/**
 * Host text as a single JSON-quoted line: newlines and quotes are escaped, so a comment body can't
 * forge a builder-owned line (a fake thread header or "Reply from …") or break out of its quotes.
 */
function quoted(text: string): string {
  return JSON.stringify(text);
}

/** The host comment id a `comment` ref points at (its typed hostRef, else its `word:comment:` id). */
export function commentIdFromRef(ref: ContextRef): string | undefined {
  if (ref.hostRef?.type === 'word.comment') return ref.hostRef.commentId;
  const prefix = 'word:comment:';
  return ref.id.startsWith(prefix) ? ref.id.slice(prefix.length) || undefined : undefined;
}

/**
 * Comment threads → attachable `comment` refs, one per thread. The typed `hostRef` carries the host
 * id the `reply <commentId> "text"` verb needs; the anchor lets a selector match the commented text.
 */
export function commentsToRefs(comments: readonly WordComment[]): ContextRef[] {
  return comments.map((c) => {
    const anchorText = c.anchorText.trim().slice(0, 120);
    return {
      id: `word:comment:${c.id}`,
      kind: 'comment',
      surface: 'word',
      title: `Comment by ${commentAuthor(c.authorName)}${c.resolved ? ' (resolved)' : ''}`,
      preview: c.content.slice(0, 120),
      ...(anchorText ? { anchor: { matchText: anchorText, locator: `comment:${c.id}` } } : {}),
      hostRef: { type: 'word.comment', commentId: c.id },
    };
  });
}

/**
 * One comment thread → a single text part: the id to reply with, its state, the commented text, the
 * comment and each reply in order. Host text is JSON-quoted onto its own line (see {@link quoted}).
 */
export function commentToContext(comment: WordComment): ResolvedContext[] {
  const anchor = comment.anchorText.trim().slice(0, MAX_COMMENT_ANCHOR_CHARS);
  const lines = [
    `Comment thread (commentId: ${comment.id}, ${comment.resolved ? 'resolved' : 'open'})`,
    ...(anchor ? [`On text: ${quoted(anchor)}`] : []),
    `${quoted(commentAuthor(comment.authorName))}: ${quoted(comment.content)}`,
    ...comment.replies.map(
      (r) => `  Reply from ${quoted(commentAuthor(r.authorName))}: ${quoted(r.content)}`,
    ),
  ];
  const [ref] = commentsToRefs([comment]);
  if (!ref) return [];
  return [{ ref, value: { as: 'text', text: lines.join('\n'), mimeType: 'text/markdown' } }];
}
