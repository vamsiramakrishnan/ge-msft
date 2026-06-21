import { z } from 'zod';

/**
 * A **content-based** anchor — the universal way to re-find a span in a host without
 * relying on per-session range GUIDs. The client resolves it with `body.search(matchText)`,
 * disambiguates with `contextHint`, and re-resolves at apply-time (see 01-architecture §7).
 *
 * One shared shape used everywhere a location is referenced: content chunks
 * (`@ge/content`), review `Finding`s, and actuation targets. Carry it; don't re-invent it.
 */
export const AnchorSchema = z.object({
  matchText: z.string(), // exact text to locate in the host
  contextHint: z.string().optional(), // disambiguates repeated matches (e.g. a section path)
  locator: z.string().optional(), // optional structural hint (e.g. "chars:120-480", "slide:4")
});
export type Anchor = z.infer<typeof AnchorSchema>;
