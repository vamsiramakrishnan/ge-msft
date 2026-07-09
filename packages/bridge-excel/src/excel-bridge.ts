import type {
  ActuationKind,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateNamedRange,
  DocStateSelection,
  DocStateSnapshot,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { buildDocStateSnapshot } from '@ge/content';
import { EXCEL_CAPABILITIES } from './capabilities.js';
import { isSet } from './capabilities-runtime.js';
import {
  rangeToContext,
  searchUsedRange,
  selectionValuesToContext,
  usedRangeToBlocks,
} from './capture.js';
import { commentAdded, deriveOrigin, documentChanged, selectionChanged } from './events.js';
import {
  formatSourceComment,
  planAddComment,
  planConditional,
  planCreateTable,
  planFormatCells,
  planInsertChart,
  planWriteCells,
  splitFormulaGrid,
} from './actuate-plan.js';
import { provenanceRecord } from './provenance-record.js';

/**
 * The exact `ActuationKind`s {@link ExcelBridge.actuate} handles (ADR-0006 closure source of
 * truth). The conformance test asserts this set equals the advertised manifest's actuation kinds.
 */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = [
  'write-cells',
  'format-cells',
  'create-table',
  'insert-chart',
  'format-conditional',
  'add-comment',
  'comment-reply',
];

/**
 * The Excel `DocBridge`. The ONLY place Office.js (`Excel.run`) is touched. Reads via the
 * native object model (selected range, used range) and maps the grid to a native table block;
 * writes via **`write-cells`** into an explicit, address-targeted range (`Sheet1!A1:B3` →
 * `worksheet.getRange`). Pure mapping lives in `capture.ts` / `actuate-plan.ts` (unit-tested);
 * this file is the host wiring.
 */
export class ExcelBridge implements DocBridge {
  readonly surface = 'excel' as const;

  /** Monotonic `<doc_state>` version, bumped on each capture (ADR-0003 Layer B element 1). */
  private docStateVersion = 0;

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
      const tables = ctx.workbook.tables;
      if (isSet('ExcelApi', '1.1')) tables.load('items/name');
      const names = ctx.workbook.names;
      if (isSet('ExcelApi', '1.7')) names.load('items/name,items/type,items/formula');
      await ctx.sync();

      const tableRanges: Array<{ name: string; range: Excel.Range }> = [];
      if (isSet('ExcelApi', '1.1')) {
        for (const table of tables.items) {
          const range = table.getRange();
          range.load('address,rowCount,columnCount');
          tableRanges.push({ name: table.name, range });
        }
        await ctx.sync();
      }

      const refs: ContextRef[] = [];
      const selValues = sel.values as string[][];
      refs.push({
        id: `xl:${sel.address}`,
        kind: 'range',
        surface: 'excel',
        title: sel.address,
        preview: hasContent(selValues) ? previewOf(selValues) : 'Blank selection',
        live: true,
      });
      for (const { name, range } of tableRanges) {
        refs.push({
          id: `xl:table:${name}`,
          kind: 'table',
          surface: 'excel',
          title: name,
          preview: `${range.address} · ${range.rowCount} × ${range.columnCount}`,
          anchor: { matchText: range.address, locator: `range:${range.address}` },
          hostRef: { type: 'excel.table', name, worksheet: parseAddress(range.address).sheetName },
        });
      }
      if (isSet('ExcelApi', '1.7')) {
        for (const named of names.items) {
          if (named.type !== 'Range' || typeof named.formula !== 'string') continue;
          const range = stripLeadingEquals(String(named.formula));
          if (!range) continue;
          refs.push({
            id: `xl:named:${named.name}`,
            kind: 'range',
            surface: 'excel',
            title: named.name,
            preview: range,
            anchor: { matchText: range, locator: `range:${range}` },
            hostRef: { type: 'excel.namedRange', name: named.name },
          });
        }
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
      const selector = ref.kind === 'selection' && ref.live ? undefined : excelSelectorFromRef(ref);
      return Excel.run(async (ctx) => {
        const sel = selector ? resolveReadRange(ctx, selector) : ctx.workbook.getSelectedRange();
        if (!sel) return [];
        sel.load('address,values,isNullObject');
        await ctx.sync();
        if ((sel as { isNullObject?: boolean }).isNullObject === true) return [];
        return selectionValuesToContext(sel.address, sel.values as string[][]);
      });
    }
    if (ref.kind === 'table') {
      const selector = excelSelectorFromRef(ref);
      if (!selector) return [];
      return Excel.run(async (ctx) => {
        const range = resolveReadRange(ctx, selector);
        if (!range) return [];
        range.load('address,values,isNullObject');
        await ctx.sync();
        if ((range as { isNullObject?: boolean }).isNullObject === true) return [];
        return rangeToContext(range.address, range.values as string[][]);
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

  canRevealContext(ref: ContextRef): boolean {
    return ref.surface === 'excel' && excelSelectorFromRef(ref) !== undefined;
  }

  async revealContext(ref: ContextRef): Promise<void> {
    if (ref.surface !== 'excel') return;
    const selector = excelSelectorFromRef(ref);
    if (!selector) return;

    await Excel.run(async (ctx) => {
      const target = resolveReadRange(ctx, selector);
      if (!target) return;
      target.load('address,isNullObject');
      await ctx.sync();
      if ((target as { isNullObject?: boolean }).isNullObject === true) return;

      const { sheetName, rangeAddress } = parseAddress(target.address);
      const sheet =
        sheetName !== undefined
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();
      sheet.activate();
      sheet.getRange(rangeAddress).select();
      await ctx.sync();
    });
  }

  /**
   * ADR-0003 Layer B element 1: an ambient structural snapshot of the workbook. Reads the active
   * sheet's used range (a table block via the shared `usedRangeToBlocks`), the current selection,
   * and the workbook's named ranges — each Office API gated on its requirement set so an older host
   * yields a partial-but-valid snapshot (we set only what we can read). Pure mapping lives in
   * `capture.ts`; this is thin host wiring (no port seam on Excel, like the rest of the bridge).
   * Version increments per capture.
   */
  async captureDocState(): Promise<DocStateSnapshot | undefined> {
    // `getUsedRangeOrNullObject` → ExcelApi 1.4 (typings l.37235): degrades on an empty sheet
    // instead of throwing. On an older host (<1.4) fall back to `getUsedRange` (1.1).
    const hasNullObj = isSet('ExcelApi', '1.4');
    // `workbook.names` is ExcelApi 1.1, but `NamedItem.formula` (the A1 reference text) is 1.7;
    // gate the named-ranges read on 1.7 so we only emit a name when we can give its range.
    const wantNames = isSet('ExcelApi', '1.7');

    const captured = await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      const used = hasNullObj ? sheet.getUsedRangeOrNullObject() : sheet.getUsedRange();
      used.load('address,values,isNullObject');

      const sel = ctx.workbook.getSelectedRange();
      sel.load('address,values');

      const names = ctx.workbook.names;
      if (wantNames) names.load('items/name,items/type,items/formula');

      await ctx.sync();

      const usedEmpty = hasNullObj && (used as { isNullObject?: boolean }).isNullObject === true;
      const usedAddress = usedEmpty ? '' : used.address;
      const usedValues = usedEmpty ? [] : (used.values as string[][]);
      const selValues = sel.values as string[][];

      const namedRanges: DocStateNamedRange[] = wantNames
        ? names.items
            .filter((n) => n.type === 'Range' && typeof n.formula === 'string')
            .map((n) => ({ name: n.name, range: stripLeadingEquals(String(n.formula)) }))
        : [];

      return {
        title: sheet.name,
        usedAddress,
        usedValues,
        selAddress: sel.address,
        selValues,
        namedRanges,
      };
    });

    const blocks = captured.usedAddress
      ? usedRangeToBlocks(captured.usedAddress, captured.usedValues)
      : [];

    this.docStateVersion += 1;
    const selection = hasContent(captured.selValues)
      ? ({
          kind: 'range',
          title: captured.selAddress,
          preview: previewOf(captured.selValues),
        } satisfies DocStateSelection)
      : undefined;

    return buildDocStateSnapshot({
      surface: 'excel',
      version: this.docStateVersion,
      title: captured.title,
      blocks,
      ...(selection ? { selection } : {}),
      ...(captured.namedRanges.length > 0 ? { namedRanges: captured.namedRanges } : {}),
    });
  }

  /**
   * ADR-0003 Layer B element 2: lazily read the workbook rows relevant to `query` instead of
   * pre-chunking the used range. Reads the active sheet's used range, then matches rows
   * (case-insensitive substring) and shapes them — header preserved — into content via the pure
   * `searchUsedRange`. Bounded; empty query / no used range / no match → `[]`.
   */
  async searchDocument(query: string): Promise<ResolvedContext[]> {
    const q = query.trim();
    if (!q) return [];
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const used = isSet('ExcelApi', '1.4')
        ? sheet.getUsedRangeOrNullObject()
        : sheet.getUsedRange();
      used.load('address,values,isNullObject');
      await ctx.sync();
      if ((used as { isNullObject?: boolean }).isNullObject === true) return [];
      return searchUsedRange(used.address, used.values as string[][], q);
    });
  }

  /**
   * ADR-0006 closure: serve the `read <A1|NamedRange>` CLI verb (ADR-0004) with an addressable,
   * on-demand read. Resolves `a1` as a sheet-qualified/plain A1 address (`Sheet1!A1:B3` → the named
   * sheet, else the active sheet) OR, when it isn't an A1 address, as a workbook-scoped **named
   * range** (`getItemOrNullObject` so a missing name degrades to `[]` instead of throwing). The
   * grid is mapped through the same pure `rangeToContext` path the rest of the bridge uses — host
   * content carried strictly as `ResolvedContext` data, never as instructions — and the read is
   * **bounded** to {@link MAX_READ_CELLS} cells so a huge range can't blow the per-turn budget.
   * Empty selector / missing name / empty grid → `[]`.
   */
  async readRange(a1: string): Promise<ResolvedContext[]> {
    const selector = a1.trim();
    if (!selector) return [];
    return Excel.run(async (ctx) => {
      const range = resolveReadRange(ctx, selector);
      if (!range) return [];
      // First sync: cheap metadata only. Bound the read BEFORE materializing `.values`, so a large
      // finite range (e.g. A1:A1048576) can't pull ~1M cells across the Office.js bridge.
      range.load('address,rowCount,columnCount,isNullObject');
      await ctx.sync();
      if ((range as { isNullObject?: boolean }).isNullObject === true) return [];
      if (!withinReadBudget(range.rowCount, range.columnCount)) return [];
      // Within budget — now materialize the values.
      range.load('values');
      await ctx.sync();
      return rangeToContext(range.address, range.values as string[][]);
    });
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    switch (req.kind) {
      case 'write-cells':
        return this.applyWriteCells(req);
      case 'format-cells':
        return this.applyFormatCells(req);
      case 'create-table':
        return this.applyCreateTable(req);
      case 'insert-chart':
        return this.applyInsertChart(req);
      case 'format-conditional':
        return this.applyConditional(req);
      case 'add-comment':
        return this.applyAddComment(req);
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
    // ADR-0003 element 3: route any `=`-prefixed cell into a formula grid so Excel evaluates an
    // inspectable, auditable formula rather than an opaque literal. Pure split (unit-tested);
    // the host write path is chosen from `hasFormulas` so non-formula writes are unchanged.
    const grid = splitFormulaGrid(plan.values);
    // Security gate (ADR-0003 §untrusted boundary): cell text is model/host-derived, so a
    // formula flagged as active-content (WEBSERVICE/DDE/external-ref/…) must never be evaluated.
    // Degrade the whole write rather than promote untrusted data into an executable instruction.
    if (grid.rejected.length > 0) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'unsafe_formula',
          message: `Refusing to evaluate ${grid.rejected.length} formula(s) flagged as unsafe (web/data/DDE/external reference).`,
        },
      };
    }
    // ADR-0003 element 4: prefer provenance.sources, fall back to params.sources, for the
    // human-visible citation comment layered on top of durable provenance metadata.
    const sources = req.provenance?.sources ?? req.params.sources ?? [];
    const result = await Excel.run(async (ctx) => {
      const { sheetName, rangeAddress } = parseAddress(target);
      const sheet =
        sheetName !== undefined
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      // Single round-trip: queue the write and the address read together. The write doesn't
      // depend on reading anything back first (the target address is supplied by the plan), so
      // there's no read-before-write ordering constraint forcing a second sync.
      if (grid.hasFormulas) {
        // Excel evaluates these; `null` cells in the formula grid are the literal cells, which
        // we set via `values` so both grids land in one batch without overwriting each other.
        range.formulas = grid.formulas as unknown[][];
        range.values = grid.values as unknown[][];
      } else {
        range.values = plan.values as unknown[][];
      }
      range.load('address');
      await ctx.sync();

      // Best-effort source comment on the anchor (first) cell. Feature-detected on ExcelApi 1.10
      // (`CommentCollection.add(string, string)`); skipped silently on older hosts so it never
      // fails the reversible, provenanced write. `comments.add` needs the full address (with
      // sheet name), so we read back the anchor cell's address before attaching.
      if (sources.length > 0 && isSet('ExcelApi', '1.10')) {
        try {
          const anchor = range.getCell(0, 0);
          anchor.load('address');
          await ctx.sync();
          ctx.workbook.comments.add(anchor.address, formatSourceComment(sources));
          await ctx.sync();
        } catch {
          // Comment attach failed (unsupported / host quirk) — the write already landed; the
          // citation comment is additive, not the system of record, so we log-and-continue.
        }
      }

      return { ok: true, changeId: req.changeId, kind: req.kind, location: range.address };
    });
    // Durable provenance (BUILD-PLAN 1.6): persist the record after the reversible write lands.
    // Best-effort, feature-detected, never fails the write (mirrors the citation-comment path).
    if (result.ok) {
      const flags = provFlags(req, await this.persistProvenance(req));
      if (Object.keys(flags).length > 0) return { ...result, ...flags };
    }
    return result;
  }

  private async applyFormatCells(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planFormatCells(req);
    if (!plan.address) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_anchor', message: 'format-cells needs target.range' },
      };
    }
    if (!plan.hasOps) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'no_format',
          message: 'format-cells needs at least one params.format field',
        },
      };
    }
    const result = await Excel.run(async (ctx) => {
      const { sheetName, rangeAddress } = parseAddress(plan.address as string);
      const sheet =
        sheetName !== undefined
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      // Each facet maps to one `range.format.*` write, applied only when present in the plan. All
      // four are ExcelApi 1.1 (typings: `RangeFont.bold` l.42865, `.italic` l.42879,
      // `RangeFill.color` l.42619, `Range.numberFormat` l.38249, `Range.format` l.38092), so no
      // per-facet requirement-set gate is needed beyond the host running Excel at all.
      if (plan.bold !== undefined) range.format.font.bold = plan.bold;
      if (plan.italic !== undefined) range.format.font.italic = plan.italic;
      if (plan.fill !== undefined) range.format.fill.color = plan.fill;
      if (plan.numberFormat !== undefined) {
        // `numberFormat` is a per-cell grid; a single code broadcasts across the whole range.
        range.numberFormat = [[plan.numberFormat]];
      }
      range.load('address');
      await ctx.sync();
      return { ok: true, changeId: req.changeId, kind: req.kind, location: range.address };
    });
    if (result.ok) {
      const flags = provFlags(req, await this.persistProvenance(req));
      if (Object.keys(flags).length > 0) return { ...result, ...flags };
    }
    return result;
  }

  /**
   * `create-table` (ADR-0007 `table` verb): promote `params.table.range` to a native Excel Table.
   * Fails closed when the range is absent. Gated on `ExcelApi 1.1` (`TableCollection.add` — typings
   * l.41116); an older host degrades rather than throwing. SECURITY (ADR-0007 §inverse-identity):
   * the recorded inverse deletes the name Excel MINTED for THIS table (`table.load('name')`, read
   * back post-sync) — never `params.table.name`, so an undo can only ever delete the object this
   * change created, not a re-resolved arbitrary table.
   */
  private async applyCreateTable(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planCreateTable(req);
    if (!plan.hasTable || !plan.address) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_anchor', message: 'create-table needs params.table.range' },
      };
    }
    // `TableCollection.add` → ExcelApi 1.1; gate so an undetectably-old host degrades, not throws.
    if (!isSet('ExcelApi', '1.1')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'unsupported_host',
          message: 'This host cannot create tables (ExcelApi < 1.1).',
        },
      };
    }
    const result = await Excel.run(async (ctx) => {
      const { sheetName } = parseAddress(plan.address as string);
      const sheet =
        sheetName !== undefined
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();
      // `tables.add(address, hasHeaders)` — pass the full (sheet-qualified) address so the table
      // lands on the intended sheet regardless of the active sheet.
      const table = sheet.tables.add(plan.address as string, plan.hasHeaders);
      const range = table.getRange();
      // Read back the MINTED name + landed address. The minted name is the inverse identity — we
      // never trust `params.table.name`. `getRange().address` gives the canonical location.
      table.load('name');
      range.load('address');
      await ctx.sync();
      return {
        ok: true as const,
        changeId: req.changeId,
        kind: req.kind,
        location: range.address,
        inverse: {
          op: 'delete-object' as const,
          objectType: 'table' as const,
          name: table.name, // the host-minted name, scoped to THIS change
        },
      };
    });
    if (result.ok) {
      const flags = provFlags(req, await this.persistProvenance(req));
      if (Object.keys(flags).length > 0) return { ...result, ...flags };
    }
    return result;
  }

  /**
   * `insert-chart` (ADR-0007 `chart` verb): add a chart over `params.chart.sourceRange` on its
   * sheet. Fails closed when the source range is absent. Gated on `ExcelApi 1.1`
   * (`ChartCollection.add` — typings l.42999); an older host degrades. SECURITY (ADR-0007
   * §inverse-identity): the inverse deletes the host-MINTED chart name (`chart.load('name')`, read
   * back post-sync), never a model-chosen label.
   */
  private async applyInsertChart(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planInsertChart(req);
    if (!plan.hasChart || !plan.address) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_anchor', message: 'insert-chart needs params.chart.sourceRange' },
      };
    }
    // `ChartCollection.add` → ExcelApi 1.1; gate so an older host degrades rather than throwing.
    if (!isSet('ExcelApi', '1.1')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'unsupported_host',
          message: 'This host cannot add charts (ExcelApi < 1.1).',
        },
      };
    }
    const result = await Excel.run(async (ctx) => {
      const { sheetName, rangeAddress } = parseAddress(plan.address as string);
      const sheet =
        sheetName !== undefined
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();
      // `charts.add(type, sourceData, seriesBy)`. The host enums are erased to strings at runtime;
      // the plan already mapped the agent enums → the `Excel.ChartType`/`ChartSeriesBy` strings.
      const chart = sheet.charts.add(
        plan.chartType as Excel.ChartType,
        sheet.getRange(rangeAddress),
        plan.seriesBy as Excel.ChartSeriesBy,
      );
      if (plan.title !== undefined) chart.title.text = plan.title;
      // Read back the MINTED chart name for the inverse identity.
      chart.load('name');
      await ctx.sync();
      return {
        ok: true as const,
        changeId: req.changeId,
        kind: req.kind,
        location: chart.name,
        inverse: {
          op: 'delete-object' as const,
          objectType: 'chart' as const,
          name: chart.name, // the host-minted name, scoped to THIS change
        },
      };
    });
    if (result.ok) {
      const flags = provFlags(req, await this.persistProvenance(req));
      if (Object.keys(flags).length > 0) return { ...result, ...flags };
    }
    return result;
  }

  /**
   * `format-conditional` (ADR-0007 `cf` verb): add one conditional-format rule to
   * `params.conditional.range`. Fails closed when the range is absent. Gated on `ExcelApi 1.6`
   * (`Range.conditionalFormats.add` — typings l.51149); an older host degrades. The rule is added
   * at the FIRST/top priority, so we read the collection count BEFORE the add and record the new
   * rule's ordinal (count-before) for the `clear-conditional` inverse — the undo clears exactly the
   * rule this change appended, by index, not by re-resolving an arbitrary rule.
   *
   * Known limitation (index-based CF addressing): if a COAUTHOR adds/removes a CF rule on the same
   * range between this change and its undo, the recorded ordinal can shift, so the undo could clear a
   * neighboring rule. This is inherent to ordinal addressing; acceptable for a best-effort inverse.
   */
  private async applyConditional(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planConditional(req);
    if (!plan.hasConditional || !plan.address || !plan.rule) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_anchor', message: 'format-conditional needs params.conditional.range' },
      };
    }
    // Security (ADR-0003 §untrusted boundary): a cellValue rule whose threshold formula is an
    // untrusted active-content vector must NOT be written — Excel evaluates a CF formula. Degrade,
    // parity with the write-cells `unsafe_formula` path; never silently evaluate it.
    if (plan.unsafe) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'unsafe_formula',
          message:
            'Refusing to write a conditional-format rule whose value is an unsafe formula (web/data/DDE/external reference).',
        },
      };
    }
    // `ConditionalFormatCollection.add` → ExcelApi 1.6; gate so an older host degrades.
    if (!isSet('ExcelApi', '1.6')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'unsupported_host',
          message: 'This host cannot add conditional formats (ExcelApi < 1.6).',
        },
      };
    }
    const rule = plan.rule;
    const result = await Excel.run(async (ctx) => {
      const { sheetName, rangeAddress } = parseAddress(plan.address as string);
      const sheet =
        sheetName !== undefined
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      // Read the rule count BEFORE the add: `add()` inserts the new rule, so its index within the
      // range's CF collection (the inverse ordinal) is the prior count.
      const cfs = range.conditionalFormats;
      const priorCount = cfs.getCount();
      range.load('address');
      await ctx.sync();
      const ruleOrdinal = priorCount.value;

      const cf = cfs.add(rule.cfType as Excel.ConditionalFormatType);
      if (rule.kind === 'cellValue') {
        cf.cellValue.rule = {
          formula1: rule.formula1,
          ...(rule.formula2 !== undefined ? { formula2: rule.formula2 } : {}),
          operator: rule.operator as Excel.ConditionalCellValueOperator,
        };
        if (rule.fill !== undefined) cf.cellValue.format.fill.color = rule.fill;
      } else if (rule.kind === 'top') {
        cf.topBottom.rule = {
          rank: rule.rank,
          type: rule.criterion as Excel.ConditionalTopBottomCriterionType,
        };
        if (rule.fill !== undefined) cf.topBottom.format.fill.color = rule.fill;
      }
      // dataBar / colorScale: the bare `add(type)` is the whole rule; nothing else to configure.
      await ctx.sync();
      return {
        ok: true as const,
        changeId: req.changeId,
        kind: req.kind,
        location: range.address,
        inverse: {
          op: 'clear-conditional' as const,
          range: range.address,
          ruleOrdinal, // index of the rule this change appended in the range's CF collection
        },
      };
    });
    if (result.ok) {
      const flags = provFlags(req, await this.persistProvenance(req));
      if (Object.keys(flags).length > 0) return { ...result, ...flags };
    }
    return result;
  }

  private async applyAddComment(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planAddComment(req);
    if (!plan.hasTarget || !plan.address) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_anchor', message: 'add-comment needs target.range' },
      };
    }
    if (!plan.hasText) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_text', message: 'add-comment needs params.text' },
      };
    }
    // `CommentCollection.add(cellAddress: string, content: string)` → ExcelApi 1.10 (typings
    // l.54943). Gate on it so an older host degrades rather than throwing. The add must target a
    // single cell, so we anchor on the first cell of `target.range` (read its full address back).
    if (!isSet('ExcelApi', '1.10')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: { code: 'unsupported', message: 'This host cannot add comments (ExcelApi < 1.10).' },
      };
    }
    const result = await Excel.run(async (ctx) => {
      const { sheetName, rangeAddress } = parseAddress(plan.address as string);
      const sheet =
        sheetName !== undefined
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();
      const anchor = sheet.getRange(rangeAddress).getCell(0, 0);
      anchor.load('address');
      // Read the anchor cell's full address back before attaching: `comments.add` needs the
      // sheet-qualified address of a single cell.
      await ctx.sync();
      ctx.workbook.comments.add(anchor.address, plan.text);
      await ctx.sync();
      return { ok: true, changeId: req.changeId, kind: req.kind, location: anchor.address };
    });
    if (result.ok) {
      const flags = provFlags(req, await this.persistProvenance(req));
      if (Object.keys(flags).length > 0) return { ...result, ...flags };
    }
    return result;
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
    const result = await Excel.run(async (ctx) => {
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
    if (result.ok) {
      const flags = provFlags(req, await this.persistProvenance(req));
      if (Object.keys(flags).length > 0) return { ...result, ...flags };
    }
    return result;
  }

  /**
   * Durable provenance persistence (BUILD-PLAN 1.6 security follow-up). After a reversible write
   * lands, stamp the {@link provenanceRecord} into the workbook's durable settings bag keyed by
   * `changeId`, so the write stays provenanced across save/reopen. Excel has no `customXmlParts`,
   * so we use `Office.context.document.settings` (workbook-persisted) carrying the JSON record.
   *
   * Best-effort and feature-detected, exactly like the citation-comment path: a missing
   * `req.provenance` is skipped (we never fabricate identity — the runtime stamps it). The reversible
   * write already succeeded; provenance is additive metadata, not the system of record, so a
   * persistence failure must NOT fail the write — but it is no longer SILENT: we return whether the
   * record dropped so the caller can flag `provenanceDropped` on the result (observability). A missing
   * settings bag (older host) is itself a drop: the change cannot be durably provenanced here.
   *
   * @returns `true` when provenance was present but could not be durably persisted; `false` when
   * persisted or when there was nothing to persist.
   */
  private async persistProvenance(req: ActuationRequest): Promise<boolean> {
    if (!req.provenance) return false; // nothing to persist — not a drop.
    try {
      const settings = (
        globalThis as {
          Office?: { context?: { document?: { settings?: OfficeSettingsLike } } };
        }
      ).Office?.context?.document?.settings;
      if (!settings) return true; // no settings bag (older host / harness) — cannot persist ⇒ drop.
      const record = provenanceRecord(req.changeId, req.provenance);
      settings.set(record.key, record.json);
      // Await the persisted save so an async settings failure is observed, not swallowed.
      const status = await new Promise<string>((resolve) => {
        try {
          settings.saveAsync((result) => resolve(asyncStatus(result)));
        } catch {
          resolve('failed');
        }
      });
      return status !== 'succeeded';
    } catch {
      return true; // settings unavailable / host quirk — the write landed, but provenance dropped.
    }
  }
}

/** The minimal slice of `Office.context.document.settings` the bridge persists provenance through. */
interface OfficeSettingsLike {
  set(name: string, value: unknown): void;
  saveAsync(callback?: (result: unknown) => void): void;
}

/** Normalize an Office `AsyncResult.status` to `'succeeded'`/`'failed'` (enum value or string). */
function asyncStatus(result: unknown): string {
  const status = (result as { status?: unknown } | undefined)?.status;
  return String(status ?? 'succeeded').toLowerCase();
}

/**
 * Observability flags for a landed write: `provenanceMissing` when the request carried no provenance
 * payload at all (an unattributed write — never mistake it for an attributed one), `provenanceDropped`
 * when a present record failed to persist durably. Persisted cleanly ⇒ empty.
 */
function provFlags(
  req: ActuationRequest,
  dropped: boolean,
): { provenanceDropped?: true; provenanceMissing?: true } {
  if (!req.provenance) return { provenanceMissing: true };
  return dropped ? { provenanceDropped: true } : {};
}

/** A NamedItem.formula is an A1 reference like "=Sheet1!$A$1:$D$9"; drop the leading `=`. */
function stripLeadingEquals(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula;
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

function excelSelectorFromRef(ref: ContextRef): string | undefined {
  const anchor = ref.anchor?.locator;
  if (anchor?.startsWith('range:')) return anchor.slice('range:'.length).trim();
  if (ref.hostRef?.type === 'excel.range') {
    return ref.hostRef.worksheet
      ? `${ref.hostRef.worksheet}!${ref.hostRef.address}`
      : ref.hostRef.address;
  }
  if (ref.hostRef?.type === 'excel.table')
    return ref.anchor?.locator?.slice('range:'.length).trim();
  if (ref.hostRef?.type === 'excel.namedRange') return ref.hostRef.name;
  if (ref.id.startsWith('xl:')) return ref.id.slice('xl:'.length).trim();
  const title = ref.title.trim();
  return title && isA1Address(title) ? title : undefined;
}

/**
 * Upper bound on cells materialized by a single `read <A1|NamedRange>` (ADR-0006). Guards the
 * per-turn budget against an unbounded selector (e.g. a whole-column "A:A"); a range over this is
 * treated as "too large to inline" and degrades to `[]`. 10k cells ≈ a generous addressable read.
 */
export const MAX_READ_CELLS = 10_000;

/**
 * True iff a `rowCount × columnCount` range is within the read budget ({@link MAX_READ_CELLS}).
 * Pure + exported so the bound that protects against a huge model-emitted `read` is unit-testable:
 * `readRange` checks this on cheap metadata BEFORE materializing `.values`. A zero/undefined
 * dimension (empty or unresolved range) is out of budget, so the port returns `[]`.
 */
export function withinReadBudget(
  rowCount: number | undefined,
  columnCount: number | undefined,
): boolean {
  const cells = (rowCount ?? 0) * (columnCount ?? 0);
  return cells > 0 && cells <= MAX_READ_CELLS;
}

/**
 * True iff `selector` looks like an A1 range reference (`A1`, `A1:B3`, `$A$1`, optionally with a
 * `Sheet1!`/`'My Sheet'!` qualifier) rather than a named range. Conservative: anything that isn't a
 * recognizable A1 shape is treated as a name, which we resolve via the workbook's named items.
 */
export function isA1Address(selector: string): boolean {
  const bang = selector.lastIndexOf('!');
  const rangePart = bang >= 0 ? selector.slice(bang + 1) : selector;
  return /^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$/.test(rangePart.trim());
}

/**
 * Resolve a read selector to a live Excel `Range`, or `undefined` when it cannot be addressed.
 * A1 selectors go through `parseAddress` + `getRange` (named sheet, else the active sheet); any
 * other selector is treated as a workbook-scoped named range via `getItemOrNullObject` (so a
 * missing name yields a null-object the caller degrades to `[]`, never a throw). Pure host wiring,
 * kept out of the bridge body so the A1-vs-name decision is unit-testable via {@link isA1Address}.
 */
function resolveReadRange(ctx: Excel.RequestContext, selector: string): Excel.Range | undefined {
  if (isA1Address(selector)) {
    const { sheetName, rangeAddress } = parseAddress(selector);
    const sheet =
      sheetName !== undefined
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
    return sheet.getRange(rangeAddress);
  }
  // Not an A1 shape → a named range. `getItemOrNullObject` degrades a missing name to a null-object
  // (its `.getRange()` is null-object-safe), which the caller filters via the `isNullObject` load.
  const named = ctx.workbook.names.getItemOrNullObject(selector);
  return named.getRange();
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
