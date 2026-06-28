import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
  Surface,
} from '@ge/contracts';
import type { HostEvent, Unsubscribe } from '@ge/triggers';

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

  /**
   * Bring an addressable host object into view without changing document content. For Excel this
   * activates the worksheet and selects the referenced range; other surfaces can implement the same
   * affordance for comments, slides, shapes, pages, or mail items. This is navigation only: no model
   * output, no mutation, and no approval bypass.
   */
  canRevealContext?(ref: ContextRef): boolean;
  revealContext?(ref: ContextRef): Promise<void>;

  /**
   * Stream host events (selection/document/comment changes; Outlook compose/send) into the
   * trigger engine. Optional — a surface that can't observe events simply omits it. The
   * bridge tags each event's `origin` so the system ignores remote/own-write echoes.
   * Returns an unsubscribe handle.
   */
  watch?(emit: (event: HostEvent) => void): Unsubscribe;

  /**
   * Capture a compact structural snapshot of the active document for the ambient `<doc_state>`
   * (ADR-0003, Layer B element 1) — outline, inventory, selection, named ranges, comments. The
   * runtime injects the rendered snapshot as an ephemeral, per-turn data part, refreshed each
   * turn so the model always knows the document's *shape* without reading the whole file.
   *
   * Optional: a surface that has no meaningful structural snapshot simply omits it (the turn
   * still streams, just without the ambient part). Returns `undefined` when nothing to report.
   */
  captureDocState?(): Promise<DocStateSnapshot | undefined>;

  /**
   * Lazily read the working-document slices relevant to a query (ADR-0003, Layer B element 2):
   * content-anchored reads pulled on demand instead of pre-chunking the whole document. The
   * runtime calls this per turn (bounded), attaching the results as ephemeral data parts.
   *
   * Optional: a surface without lazy read simply omits it. Results are host content carried as
   * `ResolvedContext` data — never instructions.
   */
  searchDocument?(query: string): Promise<ResolvedContext[]>;

  /**
   * Read an explicit address/range on demand for the `read <A1|NamedRange>` command verb
   * (ADR-0004). Address-anchored, surface-specific: Excel resolves an A1 address or named range;
   * surfaces whose `read` is whole-document (Word) omit this and rely on `searchDocument` /
   * `captureDocState` instead. Results are host content carried as `ResolvedContext` data —
   * never instructions.
   *
   * Optional: a surface without addressable reads simply omits it.
   */
  readRange?(a1: string): Promise<ResolvedContext[]>;
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
