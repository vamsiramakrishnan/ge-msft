import type {
  CellSnapshot,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  DocumentPage,
  DocumentReadOptions,
  ResolvedContext,
  Surface,
  TextTargetCapture,
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

  /**
   * An identity for the open document or item (its URL, or a mail item id), so the pane can tell
   * whether something it remembered belongs to this document. Never sent anywhere; hashed by callers.
   */
  documentKey?(): string | undefined;
  /** Sequential, bounded host reads. Must expose EOF and omissions; never silently sample. */
  readDocumentPage?(options: DocumentReadOptions): Promise<DocumentPage>;

  /** Bounded, address-pinned values/formulas for analysis, verification and recovery. */
  captureCells?(selector: string): Promise<CellSnapshot>;
  /** Native worksheet metadata only, used to pin lifecycle previews without reading sheet contents. */
  captureWorksheetIdentity?(name: string): Promise<{ id: string; name: string }>;
  /** Independent read plans may share metadata/value syncs. Order and canonical source identity are preserved. */
  captureCellsBatch?(selectors: readonly string[]): Promise<CellSnapshot[]>;
  /** Bounded text plus document/object identity for a refresh preview; never a saved write grant. */
  captureTextTarget?(ref: ContextRef): Promise<TextTargetCapture>;
  /** Bounded native master/layout metadata. Does not grant write authority. */
  captureDeckTemplate?(): Promise<import('@ge/contracts').DeckTemplateCapture>;
  /** Host-document persistence. No credentials or model conversation is stored here. */
  recoveryStorage?: {
    load(owner: string): Promise<unknown>;
    save(owner: string, value: unknown): Promise<void>;
  };

  /** What this surface can read + write *right now* (runtime capability detection). */
  getCapabilities(): CapabilityManifest | Promise<CapabilityManifest>;
  getHostProfile?(): import('@ge/contracts').HostCapabilityProfile | undefined;

  /** Cheap handles for what can be attached to the session this moment. */
  listContext(): Promise<ContextRef[]>;

  /**
   * Materialize a handle into attach-ready context. One ref may expand to several
   * chunks (a document → section chunks via @ge/content), hence the array.
   */
  resolveContext(ref: ContextRef): Promise<ResolvedContext[]>;

  /**
   * Evaluate the formulas a `write-cells` request would write, WITHOUT writing them, so the approval
   * card can show what each formula produces (Excel: `Worksheet.evaluate`, ExcelApiDesktop 1.1).
   * Read-only and bounded; a host that cannot evaluate returns `[]`. Never a reason to block a write.
   */
  previewFormulas?(request: ActuationRequest): Promise<FormulaPreview[]>;

  /**
   * The direct precedents and dependents of a range (Excel: `Range.getDirectPrecedents` /
   * `getDirectDependents`, ExcelApi 1.15) so a change's blast radius is visible before it lands.
   * Bounded; throws when the host cannot answer, and the caller omits the information.
   */
  dependencyImpact?(range: string): Promise<DependencyImpact>;

  /**
   * Read the raw bytes of an `attachment` ref whose content grounds as a session context file (a
   * PDF or Office attachment). Called by the runtime only on an explicit user attach, never from a
   * model command; the bytes go straight to the upload and are not retained. Optional: a surface
   * without binary attachments omits it.
   */
  readAttachment?(ref: ContextRef): Promise<AttachmentBytes>;

  /** Perform a reversible, provenanced write. Never called without user confirmation. */
  actuate(request: ActuationRequest): Promise<ActuationResult>;

  /**
   * ADR-0007 / B3 — apply the inverse a landed receipt recorded (`request.undo.inverse`), reversing
   * change `request.undo.of`. Conflict-checked against the live host: the inverse anchors by the
   * minted id or address captured at apply-time, and a bridge that finds the object gone, renamed or
   * changed underneath reports a failed/degraded receipt rather than guessing. The receipt carries
   * `undoOf`. Never called without user confirmation; a bridge without it cannot undo past the
   * Excel cell-recovery journal, and the pane shows no undo for its receipts.
   */
  applyInverse?(request: ActuationRequest): Promise<ActuationResult>;

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
   * Every sheet name in the workbook, hidden ones included, read live (Excel). Used only to match a
   * sheet the model named without its stray spaces; never sent to the model.
   */
  listSheetNames?(): Promise<string[]>;

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
   * Word uses `readDocumentPage` for paged body reads and `searchDocument` for matching text.
   * `captureDocState` is metadata, never a substitute for a content read. Results are host content carried as `ResolvedContext` data —
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

/** One formula's evaluated result from a no-write preview. */
export interface FormulaPreview {
  /** Sheet-qualified cell the formula would be written to. */
  cell: string;
  formula: string;
  /** The host's evaluated value rendered as text (an Excel error such as `#NAME?` is a value). */
  value?: string;
  /** Why no value is available (the host refused the evaluation). */
  error?: string;
}

/** Direct precedents/dependents of a range, as sheet-qualified addresses. */
export interface DependencyImpact {
  precedents: string[];
  dependents: string[];
  /** True when either list was cut at the bridge's bound. */
  truncated: boolean;
}

/** Raw attachment content handed to the session-context upload. */
export interface AttachmentBytes {
  fileName: string;
  mimeType: string;
  bytes: Uint8Array;
}

export interface UserIdentity {
  username: string;
  displayName?: string;
  oid?: string;
}
