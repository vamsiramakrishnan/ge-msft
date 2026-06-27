import type { Surface } from '@ge/contracts';

/**
 * Map the Office host to our `Surface`. The web-shell is surface-agnostic; the only host-
 * specific decision it makes is *which bridge to instantiate*, and that hinges on this.
 */
const HOST_TO_SURFACE: Readonly<Record<string, Surface>> = {
  Word: 'word',
  Excel: 'excel',
  PowerPoint: 'powerpoint',
  OneNote: 'onenote',
  Outlook: 'outlook',
};

/** Pure: an `Office.HostType`-like string → Surface (undefined for unknown hosts). */
export function surfaceFromHost(host: string | undefined | null): Surface | undefined {
  if (!host) return undefined;
  return HOST_TO_SURFACE[host];
}

/** Minimal shape of the Office context we read for host detection (keeps this dep-light). */
export interface OfficeContextLike {
  context?: {
    host?: unknown;
    /** Present only in Outlook. */
    mailbox?: unknown;
  };
}

/**
 * Detect the surface at runtime. Outlook exposes `mailbox` rather than a `host`, so we check
 * that first; otherwise we read `Office.context.host`. Falls back to the global `Office`.
 */
export function detectSurface(office?: OfficeContextLike): Surface | undefined {
  const o = office ?? (globalThis as { Office?: OfficeContextLike }).Office;
  const ctx = o?.context;
  if (!ctx) return undefined;
  if (ctx.mailbox) return 'outlook';
  return surfaceFromHost(ctx.host != null ? String(ctx.host) : undefined);
}
