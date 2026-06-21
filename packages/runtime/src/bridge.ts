import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
  Surface,
} from '@ge/contracts';

/**
 * The one interface every surface implements. The runtime (and the UI above it) talk
 * only to this — so Word/Excel/Outlook/PowerPoint/OneNote/Teams differ *only* in their
 * `DocBridge`. This is the seam that lets us build the loop once and reuse it everywhere
 * (see Claude's add-in: one runtime, a per-surface tool surface — same idea).
 *
 * Bridges are the ONLY code that touches Office.js / TeamsJS / Graph.
 */
export interface DocBridge {
  readonly surface: Surface;

  /** What this surface can read + write *right now* (runtime capability detection). */
  getCapabilities(): CapabilityManifest | Promise<CapabilityManifest>;

  /** Cheap handles for what can be attached to the session this moment. */
  listContext(): Promise<ContextRef[]>;

  /**
   * Materialize a handle into attach-ready context. One ref may expand to several
   * chunks (a document → section chunks via @ge/content), hence the array.
   */
  resolveContext(ref: ContextRef): Promise<ResolvedContext[]>;

  /** Perform a reversible, provenanced write. Never called without user confirmation. */
  actuate(request: ActuationRequest): Promise<ActuationResult>;
}

/**
 * The signed-in user's identity envelope. Supplies the Entra OIDC id token used as the
 * WIF subject (→ Google) and, for Plane B, a delegated Microsoft Graph token. No Google
 * credentials ever flow through here.
 */
export interface AuthClient {
  /** Entra OIDC id token whose audience the Workforce provider trusts (STS subject). */
  getIdToken(): Promise<string>;
  /** Delegated Microsoft Graph access token for the given scopes (estate plane). */
  getGraphToken?(scopes: string[]): Promise<string>;
  /** Signed-in identity for provenance stamping (e.g. "v.k@acme"). */
  getIdentity(): Promise<UserIdentity>;
}

export interface UserIdentity {
  username: string;
  displayName?: string;
  oid?: string;
}
