import { z } from 'zod';

/**
 * Branded (nominal) types for the two most safety-critical ids, so the compiler refuses to
 * mix them with each other or with arbitrary strings. We use Zod's native `.brand()` so the
 * inferred TS type carries the brand and `parse` yields a branded value — at runtime these are
 * still plain strings (the brand is erased), so all existing runtime behavior is unchanged.
 *
 * Brand at the *mint point* (where a fresh id is first created from a raw string) using the
 * helpers below; everywhere else let the branded type flow without re-branding or casting.
 */

/** Actuation idempotency / provenance key. Client-generated; makes a write idempotent. */
export const ChangeIdSchema = z.string().brand<'ChangeId'>();
export type ChangeId = z.infer<typeof ChangeIdSchema>;

/** Discovery Engine conversation (streamAssist session), used to resume a turn. */
export const SessionIdSchema = z.string().brand<'SessionId'>();
export type SessionId = z.infer<typeof SessionIdSchema>;

/** Mint a `ChangeId` from an opaque source string (the only sanctioned brand point). */
export function asChangeId(s: string): ChangeId {
  return ChangeIdSchema.parse(s);
}

/** Mint a `SessionId` from an opaque source string (the only sanctioned brand point). */
export function asSessionId(s: string): SessionId {
  return SessionIdSchema.parse(s);
}
