/**
 * Receipt enrichment for PowerPoint writes: a stable **binding** anchor, durable **tag**
 * provenance and a bounded **image preview**, all captured AFTER the write itself has synced.
 *
 * Every helper here is best effort by design: each one runs in its own `ctx.sync()` so a host that
 * rejects a binding, a tag or an image render fails only that extra step, never the landed write —
 * a queued operation that the host refuses fails the NEXT sync and takes the rest of that batch
 * with it, so the write's own sync must already be behind us when any of these are queued.
 *
 * Requirement-set versions (node_modules/@types/office-js/index.d.ts):
 *   - `Presentation.bindings` / `BindingCollection.add(shape, 'Shape', id)` /
 *     `getItemOrNullObject(id)` / `Binding.getShape()` → PowerPointApi 1.8 (l.178712, l.186895,
 *     l.186963, l.186846).
 *   - `Shape.tags` / `Slide.tags` / `TagCollection.add(key, value)` → PowerPointApi 1.3 (l.186550,
 *     l.186112, l.185983).
 *   - `Slide.getImageAsBase64({ width })` → PowerPointApi 1.8 (l.186170); `Shape.getImageAsBase64`
 *     → PowerPointApi 1.10 (l.186697). `Shape.getParentSlide()` → 1.5 (l.186704).
 */

import { provenanceRecord } from '@ge/contracts';
import type { ActuationRequest, ActuationResult } from '@ge/contracts';
import { isSet } from './capabilities-runtime.js';

/** Tag key carrying the `changeId` of the write that last touched the object. */
export const CHANGE_TAG_KEY = 'GE_CHANGE';
/** Tag key carrying the serialized provenance record (`provenanceRecord(...).json`). */
export const PROVENANCE_TAG_KEY = 'GE_PROV';
/** Prefix of every binding id the bridge mints (`binding:<id>` resolves it later). */
export const BINDING_ID_PREFIX = 'ge-';
/** Preview images are rendered at this width (pixels) — a review-card thumbnail, not an export. */
export const PREVIEW_WIDTH = 320;
/** A preview over this many decoded bytes is dropped rather than carried through the receipt. */
export const MAX_PREVIEW_BYTES = 200 * 1024;
/** A provenance record over this many characters does not fit a tag value; it is reported dropped. */
export const MAX_TAG_VALUE_CHARS = 8 * 1024;

/**
 * The deterministic binding id for a change: `ge-<changeId>` with anything outside
 * `[A-Za-z0-9_-]` folded to `-`, so the id is safe as a binding name and in a `binding:<id>`
 * selector (no `:`; the runtime splits selectors on it).
 */
export function bindingIdFor(changeId: string): string {
  return `${BINDING_ID_PREFIX}${changeId.replace(/[^A-Za-z0-9_-]/g, '-')}`;
}

/**
 * The binding id named by a shape ref, if any: `binding:<id>`, `shape:binding:<id>` or the full
 * `pp:shape:<slide>:binding:<id>` selector (the slide part is ignored — the binding knows it).
 */
export function bindingRef(shapeRef: string): string | undefined {
  const match = /^(?:(?:pp:)?shape:(?:[^:]*:)?)?binding:(.+)$/i.exec(shapeRef.trim());
  const id = match?.[1]?.trim();
  return id ? id : undefined;
}

/** What durable provenance the host did with this write. */
export type ProvenancePersistence = 'persisted' | 'dropped' | 'unsupported';

/**
 * The receipt flags for a landed write, mirroring the Excel bridge: no payload → unattributed
 * (`provenanceMissing`), a payload the host could not store → `provenanceDropped`, a host with no
 * tag support → `provenanceUnsupported`. A persisted record adds no flag.
 */
export function provenanceFlags(
  req: ActuationRequest,
  persistence: ProvenancePersistence,
): Pick<ActuationResult, 'provenanceMissing' | 'provenanceDropped' | 'provenanceUnsupported'> {
  if (!req.provenance) return { provenanceMissing: true };
  if (persistence === 'persisted') return {};
  return persistence === 'dropped' ? { provenanceDropped: true } : { provenanceUnsupported: true };
}

/** The extra receipt fields captured after a write landed. */
export type ReceiptExtras = Pick<
  ActuationResult,
  'bindingId' | 'preview' | 'provenanceMissing' | 'provenanceDropped' | 'provenanceUnsupported'
>;

/**
 * Bind `shape` to a deterministic id derived from the change (PowerPointApi 1.8). Returns the id
 * once the host confirmed it, else `undefined` (unsupported host, duplicate id, refused binding).
 */
export async function bindShape(
  ctx: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  changeId: string,
): Promise<string | undefined> {
  if (!isSet('PowerPointApi', '1.8')) return undefined;
  const id = bindingIdFor(changeId);
  try {
    ctx.presentation.bindings.add(shape, 'Shape', id);
    await ctx.sync();
    return id;
  } catch {
    return undefined;
  }
}

/** The slice of a shape or slide that carries tags. */
interface Taggable {
  readonly tags: PowerPoint.TagCollection;
}

/**
 * Stamp the change on `target` as tags (PowerPointApi 1.3): `GE_CHANGE` = changeId always, and
 * `GE_PROV` = the serialized provenance record when the request carries a payload. Reports what
 * actually persisted; a tag the host refused reads as dropped, never as recorded.
 */
export async function tagProvenance(
  ctx: PowerPoint.RequestContext,
  target: Taggable,
  req: ActuationRequest,
): Promise<ProvenancePersistence> {
  if (!isSet('PowerPointApi', '1.3')) return 'unsupported';
  const record = req.provenance ? provenanceRecord(req.changeId, req.provenance) : undefined;
  if (record && record.json.length > MAX_TAG_VALUE_CHARS) {
    // The change marker still lands; the record itself does not fit a tag value.
    try {
      target.tags.add(CHANGE_TAG_KEY, req.changeId);
      await ctx.sync();
    } catch {
      /* best effort */
    }
    return 'dropped';
  }
  try {
    target.tags.add(CHANGE_TAG_KEY, req.changeId);
    if (record) target.tags.add(PROVENANCE_TAG_KEY, record.json);
    await ctx.sync();
    return 'persisted';
  } catch {
    return 'dropped';
  }
}

/** The slice of a slide or shape that renders itself. */
interface Renderable {
  getImageAsBase64(options?: { width?: number }): OfficeExtension.ClientResult<string>;
}

/**
 * Render `source` as a PNG preview at {@link PREVIEW_WIDTH}; `undefined` when the host cannot
 * render it, the render fails, or the image exceeds {@link MAX_PREVIEW_BYTES}.
 */
export async function capturePreview(
  ctx: PowerPoint.RequestContext,
  source: Renderable,
): Promise<ActuationResult['preview'] | undefined> {
  try {
    const image = source.getImageAsBase64({ width: PREVIEW_WIDTH });
    await ctx.sync();
    const base64 = image.value;
    if (typeof base64 !== 'string' || base64.length === 0) return undefined;
    if (decodedBytes(base64) > MAX_PREVIEW_BYTES) return undefined;
    return { mimeType: 'image/png', base64, width: PREVIEW_WIDTH };
  } catch {
    return undefined;
  }
}

/** Decoded size of a base64 string, without decoding it. */
export function decodedBytes(base64: string): number {
  const padding = base64.endsWith('==') ? 2 : base64.endsWith('=') ? 1 : 0;
  return Math.floor((base64.length * 3) / 4) - padding;
}

/**
 * Everything a landed SHAPE write adds to its receipt: the binding anchor (1.8), tag provenance
 * (1.3) and a preview — the shape itself on 1.10, else its slide on 1.8. Call only after the
 * write's own sync succeeded; nothing here can fail the receipt.
 */
export async function shapeReceipt(
  ctx: PowerPoint.RequestContext,
  req: ActuationRequest,
  shape: PowerPoint.Shape,
  slide: PowerPoint.Slide,
): Promise<ReceiptExtras> {
  const bindingId = await bindShape(ctx, shape, req.changeId);
  const persistence = await tagProvenance(ctx, shape, req);
  const preview = isSet('PowerPointApi', '1.10')
    ? await capturePreview(ctx, shape)
    : isSet('PowerPointApi', '1.8')
      ? await capturePreview(ctx, slide)
      : undefined;
  return {
    ...(bindingId ? { bindingId } : {}),
    ...(preview ? { preview } : {}),
    ...provenanceFlags(req, persistence),
  };
}

/**
 * Everything a landed SLIDE write adds to its receipt: tag provenance on the slide (1.3) and a
 * slide preview (1.8). Slides have no bindings. Call only after the write's own sync succeeded.
 */
export async function slideReceipt(
  ctx: PowerPoint.RequestContext,
  req: ActuationRequest,
  slide: PowerPoint.Slide,
): Promise<ReceiptExtras> {
  const persistence = await tagProvenance(ctx, slide, req);
  const preview = isSet('PowerPointApi', '1.8') ? await capturePreview(ctx, slide) : undefined;
  return { ...(preview ? { preview } : {}), ...provenanceFlags(req, persistence) };
}

/**
 * Resolve a `binding:<id>` shape ref to the slide/shape ids it points at today (PowerPointApi
 * 1.8). The binding, not the slide ref, is authoritative: a bound shape cannot change slides, so
 * its parent slide is the right one even when the deck was re-ordered since.
 */
export async function resolveBinding(
  ctx: PowerPoint.RequestContext,
  bindingId: string,
): Promise<{ slideId: string; shapeId: string } | { error: string; gone: boolean }> {
  if (!isSet('PowerPointApi', '1.8'))
    return {
      error:
        'This PowerPoint cannot resolve binding anchors (needs PowerPointApi 1.8). ' +
        'Address the shape as pp:shape:<slide>:<shape id> instead.',
      gone: false,
    };
  const binding = ctx.presentation.bindings.getItemOrNullObject(bindingId);
  binding.load('isNullObject');
  await ctx.sync();
  if (binding.isNullObject)
    return {
      error:
        `Binding ${JSON.stringify(bindingId.slice(0, 64))} is not in this deck (the shape was ` +
        'deleted, or the deck was re-opened without it). Read the slide and use its shape id.',
      gone: true,
    };
  try {
    const shape = binding.getShape();
    const slide = shape.getParentSlide();
    shape.load('id');
    slide.load('id');
    await ctx.sync();
    return { slideId: slide.id, shapeId: shape.id };
  } catch {
    return {
      error:
        `Binding ${JSON.stringify(bindingId.slice(0, 64))} no longer points at a shape on a slide. ` +
        'Read the slide and use its shape id.',
      gone: true,
    };
  }
}
