import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { EXCEL_CAPABILITIES } from './capabilities.js';
import { isSet } from './capabilities-runtime.js';
import { rangeToContext, selectionValuesToContext } from './capture.js';
import { commentAdded, deriveOrigin, documentChanged, selectionChanged } from './events.js';
import { planWriteCells } from './actuate-plan.js';

/**
 * The Excel `DocBridge`. The ONLY place Office.js (`Excel.run`) is touched. Reads via the
 * native object model (selected range, used range) and maps the grid to a native table block;
 * writes via **`write-cells`** into an explicit, address-targeted range (`Sheet1!A1:B3` →
 * `worksheet.getRange`). Pure mapping lives in `capture.ts` / `actuate-plan.ts` (unit-tested);
 * this file is the host wiring.
 */
export class ExcelBridge implements DocBridge {
  readonly surface = 'excel' as const;

  getCapabilities(): CapabilityManifest {
    return EXCEL_CAPABILITIES;
  }

  async listContext(): Promise<ContextRef[]> {
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const sel = ctx.workbook.getSelectedRange();
      sel.load('address,values');
      const used = sheet.getUsedRange();
      used.load('address,values');
      await ctx.sync();

      const refs: ContextRef[] = [];
      const selValues = sel.values as string[][];
      if (hasContent(selValues)) {
        refs.push({
          id: `xl:${sel.address}`,
          kind: 'range',
          surface: 'excel',
          title: sel.address,
          preview: previewOf(selValues),
          live: true,
        });
      }
      refs.push({
        id: `xl:${used.address}`,
        kind: 'sheet',
        surface: 'excel',
        title: used.address,
        preview: previewOf(used.values as string[][]),
      });
      return refs;
    });
  }

  async resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    if (ref.kind === 'selection' || ref.kind === 'range') {
      return Excel.run(async (ctx) => {
        const sel = ctx.workbook.getSelectedRange();
        sel.load('address,values');
        await ctx.sync();
        return selectionValuesToContext(sel.address, sel.values as string[][]);
      });
    }
    // Sheet / table → the used range as a native table → chunks.
    return Excel.run(async (ctx) => {
      const used = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange();
      used.load('address,values');
      await ctx.sync();
      return rangeToContext(used.address, used.values as string[][]);
    });
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    switch (req.kind) {
      case 'write-cells':
        return this.applyWriteCells(req);
      case 'comment-reply':
        return this.applyCommentReply(req);
      default:
        return {
          ok: false,
          changeId: req.changeId,
          kind: req.kind,
          error: { code: 'unsupported', message: `Excel bridge cannot ${req.kind}` },
        };
    }
  }

  /**
   * Stream Excel host events into the trigger engine. Wires three Office.js events on the
   * workbook-wide collections (all sheets), maps each to a {@link HostEvent} via the pure
   * `events.ts` builders, and returns an `Unsubscribe` that removes every handler.
   *
   * - selection → `selection-changed` (no coauthor source on the args → always `'local'`).
   * - edits → `document-changed` (maps `args.source`, so a coauthor's edit is `'remote'`).
   * - comments → `comment-added`, one per added id (maps `args.source` for origin).
   *
   * Defensive: registration runs in a single `Excel.run` and feature-detects each event
   * (older requirement sets lack some). `watch` never throws — a failed registration just
   * yields an unsubscribe that removes whatever did attach. Removal happens in a fresh
   * `Excel.run`; each `EventHandlerResult.remove()` carries its own request context.
   */
  watch(emit: (event: HostEvent) => void): Unsubscribe {
    // Collected as the handlers register; the unsubscribe drains this list.
    const handles: Array<{ remove(): void }> = [];

    const registration = Excel.run(async (ctx) => {
      const workbook = ctx.workbook;
      const sheets = workbook.worksheets;

      // PRIMARY gate is the requirement-set check (`isSet`), NOT property truthiness: every
      // `onX` member is a getter that is always truthy on the Office.js proxy even when the
      // host's active set lacks the event, so `if (sheets.onChanged)` gates nothing and `.add()`
      // then THROWS on an older host. We gate each event on its confirmed ExcelApi version
      // (mapped from node_modules/@types/office-js/index.d.ts) so unsupported events simply aren't
      // registered — no throw. The per-handler try/catch stays as belt-and-suspenders, and each
      // registration is isolated so one failure can't abort the others or the whole run.

      // selection-changed — workbook-wide; args carry only `address` (no coauthor source).
      // `WorksheetCollection.onSelectionChanged` → ExcelApi 1.9 (typings l.37666).
      if (isSet('ExcelApi', '1.9')) {
        try {
          handles.push(
            sheets.onSelectionChanged.add(async (args) => {
              emit(selectionChanged(args.address));
            }),
          );
        } catch {
          // selection events not in the active requirement set — skip this one only.
        }
      }

      // document-changed — workbook-wide edits; map the coauthoring source to origin.
      // `WorksheetCollection.onChanged` → ExcelApi 1.9 (typings l.37567).
      if (isSet('ExcelApi', '1.9')) {
        try {
          handles.push(
            sheets.onChanged.add(async (args) => {
              emit(documentChanged(deriveOrigin(args.source)));
            }),
          );
        } catch {
          // change events not in the active requirement set — skip this one only.
        }
      }

      // comment-added — one HostEvent per added comment id. `args.source` gives the
      // coauthoring origin (a teammate's comment is remote); fall back to 'local'.
      // `CommentCollection.onAdded` → ExcelApi 1.12 (typings l.55031).
      const comments = workbook.comments;
      if (isSet('ExcelApi', '1.12')) {
        try {
          handles.push(
            comments.onAdded.add(async (args) => {
              const origin = deriveOrigin(args.source);
              for (const detail of args.commentDetails) {
                emit(commentAdded(detail.commentId, origin));
              }
            }),
          );
        } catch {
          // comment events not in the active requirement set — skip this one only.
        }
      }

      await ctx.sync();
    }).catch(() => {
      // Never throw from watch: a requirement-set/host failure leaves `handles` with
      // whatever attached before the error, which the unsubscribe still cleans up.
    });

    return () => {
      void registration
        .then(() =>
          Excel.run(async (ctx) => {
            for (const handle of handles) handle.remove();
            handles.length = 0;
            await ctx.sync();
          }),
        )
        .catch(() => {
          // Best-effort teardown; swallow so unsubscribe never rejects.
        });
    };
  }

  private async applyWriteCells(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planWriteCells(req);
    if (!plan.address) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_anchor', message: 'write-cells needs target.range' },
      };
    }
    if (plan.values.length === 0) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_cells', message: 'write-cells needs params.cells' },
      };
    }
    const target = plan.address;
    return Excel.run(async (ctx) => {
      const { sheetName, rangeAddress } = parseAddress(target);
      const sheet =
        sheetName !== undefined
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      // Single round-trip: queue the write and the address read together. The write doesn't
      // depend on reading anything back first (the target address is supplied by the plan), so
      // there's no read-before-write ordering constraint forcing a second sync.
      range.values = plan.values as unknown[][];
      range.load('address');
      await ctx.sync();
      return { ok: true, changeId: req.changeId, kind: req.kind, location: range.address };
    });
  }

  private async applyCommentReply(req: ActuationRequest): Promise<ActuationResult> {
    const reply = req.params.text ?? '';
    const commentId = req.params.target?.commentId;
    if (!commentId) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_comment', message: 'comment-reply needs target.commentId' },
      };
    }
    return Excel.run(async (ctx) => {
      const comments = ctx.workbook.comments;
      comments.load('items/id');
      // First sync is required: we must read back the comment ids to locate the target proxy
      // before the second (write) sync replies on it — a genuine read-then-write dependency.
      await ctx.sync();
      const comment = comments.items.find((c) => c.id === commentId);
      if (!comment) {
        return {
          ok: false,
          changeId: req.changeId,
          kind: req.kind,
          degraded: true,
          error: { code: 'comment_gone', message: 'The comment no longer exists.' },
        };
      }
      comment.replies.add(reply);
      if (req.params.resolveComment) comment.resolved = true;
      await ctx.sync();
      return { ok: true, changeId: req.changeId, kind: req.kind, location: `comment:${commentId}` };
    });
  }
}

/** True if any cell in the grid carries a non-empty value. */
function hasContent(values: string[][]): boolean {
  return values.some((row) => row.some((cell) => String(cell ?? '').trim().length > 0));
}

/** A short preview from the first row of a grid. */
function previewOf(values: string[][]): string {
  const first = values[0];
  if (!first) return '';
  return first
    .map((c) => String(c ?? ''))
    .join(' | ')
    .slice(0, 120);
}

/** Split "Sheet1!A1:B3" into its worksheet name and range address. */
export function parseAddress(address: string): { sheetName?: string; rangeAddress: string } {
  const bang = address.lastIndexOf('!');
  if (bang < 0) return { rangeAddress: address };
  const sheetName = stripSheetQuotes(address.slice(0, bang));
  const rangeAddress = address.slice(bang + 1);
  return { sheetName, rangeAddress };
}

/** Excel quotes sheet names containing spaces as 'My Sheet'; unwrap them. */
function stripSheetQuotes(name: string): string {
  if (name.startsWith("'") && name.endsWith("'") && name.length >= 2) {
    return name.slice(1, -1).replace(/''/g, "'");
  }
  return name;
}
