import type {
  Anchor,
  DocStateComment,
  DocStateInventoryEntry,
  DocStateNamedRange,
  DocStateOutlineEntry,
  DocStateSelection,
  DocStateSnapshot,
  Surface,
} from '@ge/contracts';
import type { Block } from './model.js';

/**
 * Builder + renderer for the ambient `<doc_state>` snapshot (ADR-0003, Layer B element 1).
 *
 * `buildDocStateSnapshot` derives the document outline (from heading blocks) and inventory
 * (from heading/table/slide blocks) out of the same `Block[]` the rest of the pipeline uses,
 * caps each list to a sane maximum, and flags `truncated` when it caps. `renderDocState`
 * emits the snapshot as a single compact text block **wrapped as untrusted data**.
 *
 * Security (the boundary is real, not decorative — see docs/ADR-0003 §Layer-B and CLAUDE.md):
 * every host-derived value is run through `safe()` before it is interpolated, which collapses
 * whitespace, hard-caps the length, and **escapes `& < > "`**. That escaping is what stops
 * adversarial document content (a heading or comment containing `</doc_state>`, a forged
 * `<doc_state>` re-open, or fake structural lines like `inventory:`) from breaking out of the
 * envelope and being read as instructions. Host free-text is additionally rendered **quoted**
 * so it cannot mimic the builder-emitted structure (`[kind]` tags, `label:` lines, role
 * prefixes). The only un-escaped tokens are values we control: the surface, version, ISO
 * timestamp, `ContextKind` enums, and computed dimensions.
 */

/** Caps so a pathological document can't blow the snapshot budget. */
const MAX_OUTLINE = 60;
const MAX_INVENTORY = 60;
const MAX_COMMENTS = 60;

/** How many chars of a heading's text to carry as the anchor's `matchText` prefix. */
const ANCHOR_PREFIX_CHARS = 64;

/** Per-field hard cap at render time, so one huge field can't blow the prompt budget. */
const RENDER_FIELD_CHARS = 240;

/** Code-point-safe truncation (never splits a surrogate pair mid-character). */
function clip(text: string, n: number): string {
  const cps = Array.from(text);
  return cps.length > n ? cps.slice(0, n).join('') : text;
}

/** Collapse whitespace so a value stays on a single rendered line. */
function oneLine(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

/**
 * Neutralise a host-derived string for the data envelope: single-line, length-capped (with a
 * `…` elision marker), and HTML-escaped so `<`, `>`, `&`, `"` cannot forge the `</doc_state>`
 * delimiter, a re-opening tag, or a quoted boundary. This is the load-bearing untrusted-data
 * guard — every host value MUST pass through it before interpolation.
 */
function safe(text: string, max: number = RENDER_FIELD_CHARS): string {
  const cps = Array.from(oneLine(text));
  const clipped = cps.length > max;
  const base = (clipped ? cps.slice(0, max).join('') : cps.join('')) + (clipped ? '…' : '');
  return base
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export interface BuildDocStateInput {
  surface: Surface;
  version: number;
  title?: string;
  blocks: Block[];
  selection?: DocStateSelection;
  namedRanges?: DocStateNamedRange[];
  comments?: DocStateComment[];
  /** ISO 8601 capture time; defaults to now. */
  capturedAt?: string;
  /** Injectable clock for deterministic tests. */
  now?: () => Date;
}

/**
 * Strip a leading Markdown heading marker (`## `). Native heading blocks store their text with the
 * marker (e.g. `native.heading` → `"# Service Levels"`); the outline carries clean text so the
 * renderer's own `#`.repeat(level) isn't doubled, and the anchor `matchText` matches the host's
 * real heading text (`body.search` searches document text, which has no `#`).
 */
function stripHeadingMarker(text: string): string {
  return text.replace(/^\s{0,3}#{1,6}\s+/, '');
}

/**
 * Map a heading block's `locator`/`text` to a content `Anchor` for re-finding it. The
 * `matchText` is marker-stripped + whitespace-normalised (so it matches what `body.search` sees)
 * and clipped on a code-point boundary.
 */
function headingAnchor(block: Block): Anchor {
  const matchText = clip(oneLine(stripHeadingMarker(block.text)), ANCHOR_PREFIX_CHARS);
  const anchor: Anchor = { matchText };
  if (block.locator !== undefined) anchor.locator = block.locator;
  return anchor;
}

/** Map a table block to a one-line dimension summary, when native structure is present. */
function tableSummary(block: Block): string | undefined {
  if (block.data === undefined) return undefined;
  const cols = block.data.columns.length;
  const rows = block.data.rows.length;
  return `${rows} row${rows === 1 ? '' : 's'} × ${cols} col${cols === 1 ? '' : 's'}`;
}

export function buildDocStateSnapshot(input: BuildDocStateInput): DocStateSnapshot {
  const outlineAll: DocStateOutlineEntry[] = [];
  const inventoryAll: DocStateInventoryEntry[] = [];

  let tableCount = 0;
  let slideCount = 0;

  for (const block of input.blocks) {
    if (block.kind === 'heading') {
      const headingText = stripHeadingMarker(block.text);
      const entry: DocStateOutlineEntry = {
        level: block.level ?? 0,
        text: headingText,
        anchor: headingAnchor(block),
      };
      outlineAll.push(entry);
      inventoryAll.push({
        kind: 'paragraph',
        id: block.locator ?? `heading:${outlineAll.length}`,
        title: clip(headingText, ANCHOR_PREFIX_CHARS),
      });
      continue;
    }

    if (block.kind === 'table') {
      tableCount += 1;
      const id = block.locator ?? `table:${tableCount}`;
      const inv: DocStateInventoryEntry = {
        kind: 'table',
        id,
        title: clip(block.text.split('\n')[0] ?? '', ANCHOR_PREFIX_CHARS) || `Table ${tableCount}`,
      };
      const summary = tableSummary(block);
      if (summary !== undefined) inv.summary = summary;
      inventoryAll.push(inv);
      continue;
    }

    // A slide arrives as a block whose locator marks it as a slide (e.g. "slide:4").
    if (block.locator !== undefined && block.locator.startsWith('slide:')) {
      slideCount += 1;
      inventoryAll.push({
        kind: 'slide',
        id: block.locator,
        title: clip(block.text, ANCHOR_PREFIX_CHARS) || `Slide ${slideCount}`,
      });
    }
  }

  const outline = outlineAll.slice(0, MAX_OUTLINE);
  const inventory = inventoryAll.slice(0, MAX_INVENTORY);

  const commentsAll = input.comments ?? [];
  const comments = commentsAll.slice(0, MAX_COMMENTS);

  const truncated =
    outlineAll.length > MAX_OUTLINE ||
    inventoryAll.length > MAX_INVENTORY ||
    commentsAll.length > MAX_COMMENTS;

  const capturedAt = input.capturedAt ?? (input.now ?? (() => new Date()))().toISOString();

  const snapshot: DocStateSnapshot = {
    surface: input.surface,
    version: input.version,
    capturedAt,
    outline,
    inventory,
  };
  if (input.title !== undefined) snapshot.title = input.title;
  if (input.selection !== undefined) snapshot.selection = input.selection;
  if (input.namedRanges !== undefined && input.namedRanges.length > 0) {
    snapshot.namedRanges = input.namedRanges;
  }
  if (comments.length > 0) snapshot.comments = comments;
  if (truncated) snapshot.truncated = true;

  return snapshot;
}

/**
 * Render the snapshot as a compact, untrusted-wrapped text block. The model receives this as a
 * data part; everything inside `<doc_state>…</doc_state>` is host content and must be treated as
 * data, never instructions. Host free-text is escaped via {@link safe} and quoted so it cannot
 * forge the envelope delimiter or the surrounding structure.
 */
export function renderDocState(snapshot: DocStateSnapshot): string {
  const lines: string[] = [];
  const attrs = [`surface=${snapshot.surface}`, `version=${snapshot.version}`];
  if (snapshot.truncated) attrs.push('truncated=true');
  lines.push(`<doc_state ${attrs.join(' ')}>`);
  lines.push(`capturedAt: ${snapshot.capturedAt}`);
  if (snapshot.title !== undefined) lines.push(`title: "${safe(snapshot.title)}"`);

  if (snapshot.selection !== undefined) {
    const sel = snapshot.selection;
    const preview = sel.preview !== undefined ? ` — "${safe(sel.preview)}"` : '';
    lines.push(`selection: [${sel.kind}] "${safe(sel.title)}"${preview}`);
  }

  if (snapshot.outline.length > 0) {
    lines.push('outline:');
    for (const o of snapshot.outline) {
      lines.push(`  ${'#'.repeat(Math.max(1, o.level))} "${safe(o.text)}"`);
    }
  }

  if (snapshot.inventory.length > 0) {
    lines.push('inventory:');
    for (const i of snapshot.inventory) {
      const summary = i.summary !== undefined ? ` (${safe(i.summary)})` : '';
      lines.push(`  - [${i.kind}] "${safe(i.title)}"${summary}`);
    }
  }

  if (snapshot.namedRanges !== undefined && snapshot.namedRanges.length > 0) {
    lines.push('namedRanges:');
    for (const r of snapshot.namedRanges) {
      lines.push(`  - "${safe(r.name)}" = "${safe(r.range)}"`);
    }
  }

  if (snapshot.comments !== undefined && snapshot.comments.length > 0) {
    lines.push('comments:');
    for (const c of snapshot.comments) {
      const who = c.author !== undefined ? `"${safe(c.author)}": ` : '';
      const hint = c.anchorHint !== undefined ? ` @"${safe(c.anchorHint)}"` : '';
      lines.push(`  - ${who}"${safe(c.text)}"${hint}`);
    }
  }

  lines.push('</doc_state>');
  return lines.join('\n');
}
