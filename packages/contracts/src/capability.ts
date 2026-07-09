import { z } from 'zod';
import { ChangeIdSchema } from './brand.js';
import { ContextKindSchema, SurfaceSchema, type Surface } from './context.js';
import { SourceRefSchema } from './finding.js';
import { ProvenancePayloadSchema } from './provenance.js';

/**
 * The capability foundation (part 2 of 2): **actuation**.
 *
 * Everything a surface can *write* is advertised as an `Actuation` and invoked via an
 * `ActuationRequest`. Writes are reversible and provenanced by contract: each carries a
 * client-generated `changeId` (idempotent re-apply) and an optional `ProvenancePayload`
 * stamped into the host's durable metadata. Experiences/agents compose these; they never
 * touch Office.js directly. See docs/ADR-0002-capability-model.md.
 */

/**
 * The COMPLETE cross-surface actuation catalog (ADR-0007 §cross-surface; the authoritative map is
 * `docs/CAPABILITY-CATALOG.md`). Every kind here is grounded in the host typings (`@types/office-js`)
 * or a named Graph endpoint. The enum is the *locked map* — a kind appears here once modeled, but is
 * only **advertised** by a surface once its bridge HANDLES it (ADR-0006 closure), so unimplemented
 * kinds sit in the catalog without breaking conformance (exactly as `create-event`/`post-card` did).
 *
 * **Plane A** = client-direct host writes (Office.js / TeamsJS). **Plane B** = estate writes via
 * Microsoft Graph — they need delegated `*.ReadWrite`/`*.Send` scopes + a `GraphClient` write path
 * that do not exist yet, so they are catalogued but gated behind the estate-auth path.
 */
export const ActuationKindSchema = z.enum([
  // ── Surface-agnostic content writes ───────────────────────────────────────
  'insert-text', // insert plain text at the cursor/selection
  'replace-selection', // replace the current selection
  'insert-ooxml', // Word/PPT: insert rich OOXML
  'insert-html', // Word/OneNote: insert a screened rich-HTML fragment
  'insert-image', // Word/PPT/OneNote: insert a base64 image
  'insert-table', // Word/PPT/OneNote: build a native table from a value grid
  'insert-hyperlink', // Word/PPT/OneNote: add/anchor a hyperlink
  'add-comment', // Word/Excel: add a new content/cell-anchored comment (ADR-0004 `comment` verb)
  'comment-reply', // Word/Excel/PPT: reply to (and optionally resolve) a comment
  // ── Word ──────────────────────────────────────────────────────────────────
  'tracked-change', // Word: insert/replace as a tracked change
  'fill-content-control', // Word: populate a named content control
  'insert-content-control', // Word: insert a NEW content control
  'insert-paragraph', // Word: insert a structural paragraph block
  'apply-style', // Word: apply a named (built-in/custom) style
  'define-style', // Word: define a new document style (WordApi 1.5; facets desktop-only)
  'apply-list', // Word: apply numbered/bulleted list formatting
  'insert-bookmark', // Word: insert a named bookmark
  'insert-field', // Word: insert a field (TOC / cross-ref / page number / date)
  'insert-break', // Word: insert a page/section/line break
  'set-page-layout', // Word: margins/orientation/size (WordApiDesktop — degrades on web)
  'set-header-footer', // Word: write a section header/footer body
  'insert-note', // Word: insert a footnote/endnote
  'resolve-revisions', // Word: accept/reject tracked revisions (IRREVERSIBLE — reversible:false)
  'find-replace', // Word: search + replace across the body (multi-range blast radius)
  'set-doc-properties', // Word: set built-in/custom document properties
  // ── Excel ───────────────────────────────────────────────────────────────
  'write-cells', // Excel: write values/formulas to a range
  'format-cells', // Excel: apply formatting (bold/fill/numberFormat) to a range (ADR-0004 `format`)
  'create-table', // Excel: promote a range to a native Table (ADR-0007 `table` verb)
  'insert-chart', // Excel: add a chart over a source range (ADR-0007 `chart` verb)
  'format-conditional', // Excel: add a conditional-format rule to a range (ADR-0007 `cf` verb)
  'insert-pivot', // Excel: add a native PivotTable over a source range
  'sort-range', // Excel: apply native range/table sort
  'filter-range', // Excel: apply native AutoFilter/table filter criteria
  'manage-worksheet', // Excel: create/rename/delete/activate/protect a worksheet
  'format-chart', // Excel: update chart title/legend/axes/style without recreating it
  'set-entity-card', // Excel: attach a linked-entity card to a cell (typings-limited)
  // ── PowerPoint ────────────────────────────────────────────────────────────
  'insert-slide', // PowerPoint: add a slide
  'set-speaker-notes', // PowerPoint: set a slide's speaker notes (no host write path — unadvertised)
  'add-shape', // PowerPoint: add a text box / geometric shape / connector line
  'add-table-slide', // PowerPoint: add a native table (seeded values) onto a slide
  'set-shape-text', // PowerPoint: replace the text of an existing shape
  'format-shape', // PowerPoint: fill/line/font/geometry of a shape
  'delete-slide', // PowerPoint: delete a slide (snapshot-undo via exportAsBase64)
  'move-slide', // PowerPoint: reorder a slide
  'duplicate-slide', // PowerPoint: duplicate a slide
  'apply-slide-layout', // PowerPoint: apply an existing layout
  // ── OneNote (web-only, legacy manifest) ───────────────────────────────────
  'append-page', // OneNote: add a synthesized page
  'add-outline', // OneNote: add an HTML outline onto the active page
  'append-html', // OneNote: append HTML to an outline (coarse undo — degrade)
  'append-rich-text', // OneNote: append a plain paragraph
  'create-section', // OneNote: create a section (no in-API delete → estate undo)
  'create-section-group', // OneNote: create a section group
  'set-page-title', // OneNote: rename the page (fully reversible — restore-text)
  'add-note-tag', // OneNote: add a To-Do/Important/Question tag
  'insert-html-at-cursor', // OneNote: insert HTML at the cursor (no anchor/undo — last resort)
  'insert-link-at-cursor', // OneNote: embed a link at the cursor
  // ── Outlook (Plane A — compose; draft-reviewable via the on-send gate) ─────
  'create-mail', // Outlook: open a draft message
  'reply-mail', // Outlook: open a grounded reply
  'add-attachment', // Outlook: attach a file (url/base64) or item to the draft
  'remove-attachment', // Outlook: remove an attachment from the draft
  'set-recipients', // Outlook: set/add to/cc/bcc
  'set-subject', // Outlook: set the subject
  'set-body', // Outlook: replace the body
  'prepend-body', // Outlook: prepend to the body
  'add-categories', // Outlook: add/remove categories (immediate on a received item)
  'set-sensitivity-label', // Outlook: set the MIP sensitivity label (Mailbox 1.13)
  'set-sensitivity', // Outlook: set the legacy sensitivity class (Mailbox 1.14)
  'set-internet-headers', // Outlook: set custom x-* headers (provenance carrier)
  'add-notification', // Outlook: in-pane notification (UI affordance — not gated)
  'compose-appointment', // Outlook: open a new appointment/meeting form
  'set-delay-delivery', // Outlook: set delayed delivery time (Mailbox 1.13)
  'save-draft', // Outlook: persist the active draft
  // ── Teams (Plane A — staged/reviewable) ───────────────────────────────────
  'post-card', // Teams: post an Adaptive Card (notes/action items)
  'post-message', // Teams: stage a reviewable chat post / Adaptive Card (reversible)
  // ── Estate (Plane B — Microsoft Graph; needs estate-auth OAuth scopes) ────
  'create-event', // Outlook/Graph: create a calendar item (Calendars.ReadWrite)
  'create-task', // Planner/To Do/Graph: create a task (Tasks.ReadWrite)
  'move-message', // Outlook/Graph: move a message to a folder (Mail.ReadWrite)
  'copy-message', // Outlook/Graph: copy a message to a folder
  'categorize-message', // Outlook/Graph: categorize/triage a received message
  'flag-message', // Outlook/Graph: set follow-up flag (Graph-only)
  'delete-message', // Outlook/Graph: soft-delete a message
  'create-mail-rule', // Outlook/Graph: create an inbox rule (high-privilege; MailboxSettings.ReadWrite)
  'post-chat-message', // Teams/Graph: post a 1:1/group chat message (ChatMessage.Send)
  'post-channel-message', // Teams/Graph: post a channel message (ChannelMessage.Send)
  'reply-channel-message', // Teams/Graph: reply in a channel thread
  'update-message', // Teams/Graph: update a posted message/card
  'set-reaction', // Teams/Graph: set/unset an emoji reaction (reversible)
  'create-online-meeting', // Teams/Graph: create an online meeting (OnlineMeetings.ReadWrite)
  'send-activity-notification', // Teams/Graph: proactive activity-feed notify (IRREVERSIBLE; TeamsActivity.Send)
  'graph-create-page', // OneNote/Graph: create a page server-side (Notes.Create)
  'graph-patch-page', // OneNote/Graph: targeted append/insert/replace on a page (Notes.ReadWrite)
  'graph-create-section', // OneNote/Graph: create a section server-side
]);
export type ActuationKind = z.infer<typeof ActuationKindSchema>;

/** What a surface advertises it can do (drives the UI's available actions). */
export const ActuationSchema = z.object({
  kind: ActuationKindSchema,
  surface: SurfaceSchema,
  title: z.string(), // verb shown to the user ("Insert as tracked change")
  description: z.string().optional(),
  reversible: z.boolean(), // false ⇒ the UI must warn before invoking
  appliesTo: z.array(ContextKindSchema).optional(), // context kinds this can target
});
export type Actuation = z.infer<typeof ActuationSchema>;

/** Parameters for an actuation. Open by design — agents fill what a kind needs. */
export const ActuationParamsSchema = z.object({
  text: z.string().optional(),
  ooxml: z.string().optional(),
  html: z.string().optional(),
  /** A content-anchored target (matchText/contextHint) or an explicit host id. */
  target: z
    .object({
      matchText: z.string().optional(),
      contextHint: z.string().optional(),
      contentControlId: z.string().optional(),
      commentId: z.string().optional(),
      range: z.string().optional(),
      slideIndex: z.number().optional(),
      slideId: z.string().optional(),
      // ADR-0007 cross-surface anchors: PowerPoint shapes, OneNote outline/paragraph/region,
      // and Plane-B Graph message ids. All optional — a kind reads the anchor it needs.
      shapeId: z.string().optional(),
      shapeName: z.string().optional(),
      outlineId: z.string().optional(),
      paragraphId: z.string().optional(),
      pageContentId: z.string().optional(),
    })
    .optional(),
  /** Shared insert position for structural inserts (Word breaks/paragraphs, OneNote/PPT inserts). */
  position: z.enum(['start', 'end', 'before', 'after', 'replace']).optional(),
  cells: z.array(z.array(z.string())).optional(), // write-cells
  /** format-cells: host-native formatting applied to `target.range` (ADR-0004 `format` verb). */
  format: z
    .object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      fill: z.string().optional(), // background color, e.g. "#FFF2CC"
      numberFormat: z.string().optional(), // e.g. "$#,##0.00"
    })
    .optional(),
  /** create-table (ADR-0007): promote `range` to a native Table. */
  table: z
    .object({
      range: z.string(),
      hasHeaders: z.boolean().default(true),
      name: z.string().optional(), // table name; the bridge mints one if absent
    })
    .optional(),
  /** insert-chart (ADR-0007): a chart over `sourceRange`. */
  chart: z
    .object({
      chartType: z.enum(['column', 'bar', 'line', 'pie', 'scatter', 'area']),
      sourceRange: z.string(),
      seriesBy: z.enum(['rows', 'columns', 'auto']).default('auto'),
      title: z.string().optional(),
    })
    .optional(),
  /** format-conditional (ADR-0007): one conditional-format rule applied to `range`. */
  conditional: z
    .object({
      range: z.string(),
      rule: z.discriminatedUnion('kind', [
        z.object({
          kind: z.literal('cellValue'),
          operator: z.enum(['gt', 'lt', 'ge', 'le', 'eq', 'ne', 'between']),
          value: z.string(),
          value2: z.string().optional(), // upper bound for `between`
          fill: z.string().optional(), // highlight color, e.g. "#C6EFCE"
        }),
        z.object({ kind: z.literal('dataBar') }),
        z.object({ kind: z.literal('colorScale') }),
        z.object({
          kind: z.literal('top'),
          rank: z.number(), // top/bottom N
          bottom: z.boolean().default(false),
          fill: z.string().optional(),
        }),
      ]),
    })
    .optional(),
  /** insert-pivot (Excel): native PivotTable creation over an explicit source range. */
  pivot: z
    .object({
      sourceRange: z.string(),
      destinationRange: z.string(),
      name: z.string().optional(),
      rowFields: z.array(z.string()).default([]),
      columnFields: z.array(z.string()).default([]),
      valueFields: z.array(z.string()).default([]),
      filterFields: z.array(z.string()).default([]),
    })
    .optional(),
  /** sort-range (Excel): native sort over a range/table. */
  sortRange: z
    .object({
      range: z.string(),
      key: z.string(),
      order: z.enum(['ascending', 'descending']).default('ascending'),
      hasHeaders: z.boolean().default(true),
    })
    .optional(),
  /** filter-range (Excel): native filter over a range/table column. */
  filterRange: z
    .object({
      range: z.string(),
      column: z.string(),
      criterion1: z.string(),
      operator: z.string().optional(),
      criterion2: z.string().optional(),
    })
    .optional(),
  /** manage-worksheet (Excel): worksheet lifecycle and protection operations. */
  worksheet: z
    .object({
      action: z.enum(['create', 'rename', 'delete', 'activate', 'protect', 'unprotect']),
      name: z.string(),
      newName: z.string().optional(),
      position: z.enum(['before', 'after', 'end']).optional(),
    })
    .optional(),
  /** format-chart (Excel): update an existing chart without recreating its data source. */
  chartFormat: z
    .object({
      chartId: z.string().optional(),
      chartName: z.string().optional(),
      title: z.string().optional(),
      legend: z.enum(['show', 'hide']).optional(),
      style: z.string().optional(),
      xAxisTitle: z.string().optional(),
      yAxisTitle: z.string().optional(),
    })
    .optional(),
  slide: z
    .object({ title: z.string(), bullets: z.array(z.string()), notes: z.string().optional() })
    .optional(),
  /** PowerPoint prebuilt deck import: a generated PPTX artifact inserted atomically. */
  deck: z
    .object({
      format: z.literal('pptx').default('pptx'),
      base64: z.string().min(1),
      slideCount: z.number().int().positive().optional(),
      formatting: z.enum(['KeepSourceFormatting', 'UseDestinationTheme']).optional(),
      targetSlideId: z.string().optional(),
      specFingerprint: z.string().optional(),
    })
    .optional(),
  mail: z
    .object({
      to: z.array(z.string()).optional(),
      cc: z.array(z.string()).optional(),
      bcc: z.array(z.string()).optional(),
      subject: z.string().optional(),
      body: z.string().optional(),
      recipientMode: z.enum(['set', 'add']).optional(), // setAsync replaces, addAsync appends
      coercion: z.enum(['html', 'text']).optional(),
    })
    .optional(),
  // ─── ADR-0007 cross-surface params (the locked capability catalog) ─────────
  // All optional/additive — a bridge reads only the field(s) its kind needs. Grounded in
  // docs/CAPABILITY-CATALOG.md; see that file for the per-field API + requirement set.

  /** insert-image (Word/PPT/OneNote): a base64 image + optional geometry. */
  image: z
    .object({
      base64: z.string(),
      altText: z.string().optional(),
      width: z.number().optional(),
      height: z.number().optional(),
      left: z.number().optional(),
      top: z.number().optional(),
    })
    .optional(),
  /** insert-table (Word/PPT/OneNote): a native table built from a value grid (vs Excel `table`). */
  tableGrid: z
    .object({
      rows: z.array(z.array(z.string())),
      hasHeaders: z.boolean().default(false),
      left: z.number().optional(),
      top: z.number().optional(),
      width: z.number().optional(),
      height: z.number().optional(),
    })
    .optional(),
  /** insert-hyperlink (Word/PPT/OneNote). The URL is untrusted — screen the scheme/host. */
  hyperlink: z
    .object({
      url: z.string(),
      displayText: z.string().optional(),
      screenTip: z.string().optional(),
    })
    .optional(),
  /** apply-style / define-style (Word). */
  style: z
    .object({
      name: z.string(),
      builtIn: z.boolean().optional(),
      define: z
        .object({
          type: z.enum(['paragraph', 'character', 'table', 'list']),
          basedOn: z.string().optional(),
        })
        .optional(),
    })
    .optional(),
  /** apply-list (Word). */
  list: z
    .object({ kind: z.enum(['numbered', 'bulleted']), level: z.number().optional() })
    .optional(),
  /** insert-bookmark (Word). */
  bookmark: z.object({ name: z.string() }).optional(),
  /** insert-field (Word): TOC / cross-reference / page number / date. */
  field: z.object({ fieldType: z.string(), code: z.string().optional() }).optional(),
  /** insert-content-control (Word): a NEW content control (vs fill-content-control). */
  contentControl: z
    .object({ type: z.string(), tag: z.string().optional(), title: z.string().optional() })
    .optional(),
  /** insert-break (Word). */
  break: z
    .object({ breakType: z.enum(['page', 'sectionNext', 'sectionContinuous', 'line']) })
    .optional(),
  /** set-page-layout (Word; WordApiDesktop — degrades on the web). */
  pageLayout: z
    .object({
      orientation: z.enum(['portrait', 'landscape']).optional(),
      margins: z
        .object({
          top: z.number().optional(),
          bottom: z.number().optional(),
          left: z.number().optional(),
          right: z.number().optional(),
        })
        .optional(),
      size: z.string().optional(),
    })
    .optional(),
  /** set-header-footer (Word). */
  headerFooter: z
    .object({
      which: z.enum(['header', 'footer']),
      type: z.enum(['primary', 'firstPage', 'evenPages']).default('primary'),
      text: z.string().optional(),
      html: z.string().optional(),
    })
    .optional(),
  /** insert-note (Word): footnote/endnote. */
  note: z.object({ kind: z.enum(['footnote', 'endnote']), text: z.string() }).optional(),
  /** resolve-revisions (Word): accept/reject tracked changes — IRREVERSIBLE. */
  revisions: z
    .object({
      scope: z.enum(['one', 'all', 'shown']),
      action: z.enum(['accept', 'reject']),
      trackedChangeId: z.string().optional(),
    })
    .optional(),
  /** find-replace (Word): preview must show the hit count before applying. */
  findReplace: z
    .object({
      find: z.string(),
      replace: z.string(),
      matchCase: z.boolean().optional(),
      matchWholeWord: z.boolean().optional(),
    })
    .optional(),
  /** set-doc-properties (Word): built-in + custom document properties. */
  docProperties: z.record(z.string()).optional(),
  /** add-shape (PowerPoint): text box / geometric shape / connector line. */
  shape: z
    .object({
      shapeType: z.enum(['textBox', 'geometric', 'line']),
      geometryType: z.string().optional(), // geometric: Rectangle/Ellipse/Chevron/…
      connectorType: z.enum(['straight', 'elbow', 'curve']).optional(), // line
      text: z.string().optional(),
      fill: z.string().optional(),
      left: z.number().optional(),
      top: z.number().optional(),
      width: z.number().optional(),
      height: z.number().optional(),
    })
    .optional(),
  /** format-shape (PowerPoint): in-place shape formatting (snapshot-prior inverse). */
  shapeFormat: z
    .object({
      fill: z.string().optional(),
      line: z.string().optional(),
      font: z
        .object({
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
          underline: z.boolean().optional(),
          color: z.string().optional(),
          size: z.number().optional(),
          name: z.string().optional(),
        })
        .optional(),
      zOrder: z.enum(['front', 'back', 'forward', 'backward']).optional(),
    })
    .optional(),
  /** slide management (PowerPoint): move/duplicate/delete by index or id. */
  slideOp: z
    .object({
      fromIndex: z.number().optional(),
      toIndex: z.number().optional(),
      slideId: z.string().optional(),
    })
    .optional(),
  /** apply-slide-layout (PowerPoint). */
  layout: z
    .object({ layoutId: z.string().optional(), layoutName: z.string().optional() })
    .optional(),
  /** add-note-tag (OneNote): the native action-item primitive. */
  noteTag: z.object({ type: z.string(), status: z.string().default('unknown') }).optional(),
  /** create-section / create-section-group (OneNote). */
  section: z
    .object({
      name: z.string(),
      parent: z.enum(['notebook', 'group']).optional(),
      insertLocation: z.enum(['before', 'after']).optional(),
    })
    .optional(),
  /** set-page-title (OneNote): the one fully-reversible OneNote write. */
  pageTitle: z.string().optional(),
  /** add-attachment / remove-attachment (Outlook). The URI is untrusted — allowlist it. */
  attachment: z
    .object({
      uri: z.string().optional(),
      base64: z.string().optional(),
      itemId: z.string().optional(),
      attachmentId: z.string().optional(),
      name: z.string().optional(),
      isInline: z.boolean().optional(),
    })
    .optional(),
  /** add-categories (Outlook): names must pre-exist in the master-categories list. */
  categories: z
    .object({ add: z.array(z.string()).optional(), remove: z.array(z.string()).optional() })
    .optional(),
  /** set-sensitivity-label (labelId) / set-sensitivity (class) (Outlook). */
  sensitivity: z
    .object({
      labelId: z.string().optional(),
      class: z.enum(['normal', 'personal', 'private', 'confidential']).optional(),
    })
    .optional(),
  /** add-notification (Outlook): an in-pane status affordance, not a mail mutation. */
  notification: z
    .object({
      key: z.string(),
      type: z.enum(['informationalMessage', 'progressIndicator', 'errorMessage', 'insightMessage']),
      message: z.string(),
      persistent: z.boolean().optional(),
    })
    .optional(),
  /** set-internet-headers (Outlook): custom x-* headers only (provenance carrier). */
  headers: z
    .object({ set: z.record(z.string()).optional(), remove: z.array(z.string()).optional() })
    .optional(),
  /** set-delay-delivery (Outlook): ISO datetime. */
  delayDeliveryUntil: z.string().optional(),
  /** compose-appointment (Outlook, Plane A) / create-event (Graph, Plane B). */
  appointment: z
    .object({
      subject: z.string().optional(),
      start: z.string().optional(),
      end: z.string().optional(),
      location: z.string().optional(),
      requiredAttendees: z.array(z.string()).optional(),
      optionalAttendees: z.array(z.string()).optional(),
      body: z.string().optional(),
      isOnlineMeeting: z.boolean().optional(),
    })
    .optional(),
  /** Plane B (Outlook/Graph): message triage target + fields (move/categorize/flag/delete). */
  message: z
    .object({
      id: z.string(),
      destinationFolderId: z.string().optional(),
      importance: z.enum(['low', 'normal', 'high']).optional(),
      isRead: z.boolean().optional(),
      flagStatus: z.enum(['flagged', 'complete', 'notFlagged']).optional(),
    })
    .optional(),
  /**
   * create-mail-rule (Outlook/Graph): high-privilege standing automation — the classic mailbox
   * auto-exfiltration vector. UNTRUSTED SINK: `actions` forward/redirect recipients must be
   * ALLOWLIST-screened, not merely user-confirmed (a forwarding rule persists and leaks future mail);
   * never derive rule actions from email-body content. Hard-gated AND screened.
   */
  mailRule: z
    .object({
      displayName: z.string(),
      conditions: z.unknown(),
      actions: z.unknown(),
      sequence: z.number().optional(),
    })
    .optional(),
  /** Plane B (Teams/Graph): the post/notify/react target ids. */
  graphTarget: z
    .object({
      chatId: z.string().optional(),
      teamId: z.string().optional(),
      channelId: z.string().optional(),
      messageId: z.string().optional(),
      userId: z.string().optional(),
    })
    .optional(),
  /**
   * post-card / update-message (Teams/Graph): a typed Adaptive Card (vs smuggling via `html`).
   * UNTRUSTED SINK: card `Action.OpenUrl`/`Action.Submit` targets and input fields are
   * model/host-derived — screen them (allowlist scheme/host) at apply-time exactly like `hyperlink`.
   * Gating (user confirm) is NOT screening; a confirmed card can still carry a malicious action url.
   */
  card: z.object({ adaptiveCard: z.unknown(), fallbackText: z.string().optional() }).optional(),
  /** set-reaction (Teams/Graph). */
  reaction: z.object({ reactionType: z.string(), emoji: z.string().optional() }).optional(),
  /** create-online-meeting (Teams/Graph). */
  onlineMeeting: z
    .object({
      subject: z.string().optional(),
      startDateTime: z.string().optional(),
      endDateTime: z.string().optional(),
      participants: z.array(z.string()).optional(),
    })
    .optional(),
  /** send-activity-notification (Teams/Graph): IRREVERSIBLE; activityType must be in the manifest. */
  activity: z
    .object({
      activityType: z.string(),
      previewText: z.string(),
      recipientUserId: z.string(),
      topicValue: z.string().optional(),
      templateParameters: z.record(z.string()).optional(),
    })
    .optional(),
  /** create-task (Planner/To Do, Graph). */
  task: z
    .object({
      title: z.string(),
      planId: z.string().optional(),
      bucketId: z.string().optional(),
      listId: z.string().optional(),
      assignees: z.array(z.string()).optional(),
      dueDateTime: z.string().optional(),
    })
    .optional(),
  resolveComment: z.boolean().optional(),
  sources: z.array(SourceRefSchema).optional(),
});
export type ActuationParams = z.infer<typeof ActuationParamsSchema>;

export const ActuationRequestSchema = z.object({
  changeId: ChangeIdSchema, // client-generated; makes the write idempotent
  kind: ActuationKindSchema,
  surface: SurfaceSchema,
  params: ActuationParamsSchema,
  provenance: ProvenancePayloadSchema.optional(),
});
export type ActuationRequest = z.infer<typeof ActuationRequestSchema>;

/**
 * The recorded INVERSE of an actuation (ADR-0007). Reversibility is an explicit, recorded operation
 * rather than an implicit property of a tracked change: when a bridge lands a write it reports HOW to
 * undo it. Pure additions (table/chart/pivot) carry a `delete-object` descriptor keyed by the object
 * name the host minted; in-place mutations (conditional formatting, a grid spilled over existing data)
 * carry the prior state needed to restore it. The descriptor is persisted alongside provenance so an
 * undo is auditable and does not depend on host-session undo state.
 */
export const InverseDescriptorSchema = z.discriminatedUnion('op', [
  z.object({
    op: z.literal('delete-object'),
    // The host object kind to delete. SECURITY (ADR-0007 §inverse-identity): `name` MUST be the
    // apply-time MINTED handle/range-id recorded for THIS change (scoped to its `changeId`/provenance
    // entry) — NEVER a human label or a re-resolvable search string. The undo applier verifies the
    // object still matches that minted id and DEGRADES on a mismatch; it must never re-resolve an
    // arbitrary object of this type against the live host (or an undo could delete a hand-made
    // table/chart). Word objects are mostly unnamed, so the recorded id is the inserted-range id.
    // The Phase-B+ security-reviewer pass asserts this per bridge.
    objectType: z.enum([
      'table',
      'chart',
      'pivot', // Excel
      'shape',
      'slide', // PowerPoint
      'bookmark',
      'field',
      'note',
      'style',
      'paragraph',
      'inline-picture',
      'content-control',
      'word-table',
      'hyperlink', // Word
      'outline-region',
      'page-content',
      'note-tag', // OneNote
      'attachment',
      'category', // Outlook (Plane A)
      'chat-message',
      'channel-message',
      'reaction', // Teams (Plane B)
      'event',
      'online-meeting',
      'task',
      'message',
      'mail-rule', // Plane B
      'onenote-page',
      'onenote-section', // OneNote (Plane B)
    ]),
    name: z.string(), // the host object name/id to delete
  }),
  z.object({
    op: z.literal('restore-values'),
    range: z.string(),
    values: z.array(z.array(z.string())), // prior cell values to write back
  }),
  z.object({
    op: z.literal('clear-conditional'),
    range: z.string(),
    ruleOrdinal: z.number(), // index of the added rule within the range's CF collection
  }),
  // ── ADR-0007 cross-surface inverses (Word/PPT/OneNote/Outlook/Teams) ───────
  // In-place mutations record the prior state needed to restore it; each anchors by a re-resolvable
  // id captured at apply-time, per the inverse-identity rule above.
  z.object({ op: z.literal('restore-text'), anchor: z.string(), priorText: z.string() }), // header/footer, find-replace, replace-selection, PPT shape text, OneNote title
  z.object({ op: z.literal('restore-style'), anchor: z.string(), priorStyle: z.string() }), // Word apply-style
  z.object({
    op: z.literal('restore-content-control'),
    contentControlId: z.string(),
    priorText: z.string(),
  }), // Word fill-content-control
  z.object({ op: z.literal('restore-doc-properties'), prior: z.record(z.string()) }), // Word set-doc-properties
  z.object({ op: z.literal('detach-list'), anchor: z.string() }), // Word apply-list
  z.object({
    op: z.literal('restore-slide'),
    base64: z.string(),
    targetSlideId: z.string().optional(),
  }), // PPT delete-slide undo
  z.object({ op: z.literal('move-slide'), toIndex: z.number() }), // PPT move-slide undo
  z.object({ op: z.literal('apply-layout'), layoutId: z.string() }), // PPT apply-slide-layout undo
  z.object({
    op: z.literal('restore-shape-format'),
    shapeId: z.string(),
    prior: z.record(z.string()),
  }), // PPT format-shape
  z.object({ op: z.literal('restore-message'), messageId: z.string(), priorContent: z.string() }), // Teams update-message
  z.object({
    op: z.literal('restore-mail-state'),
    messageId: z.string(),
    prior: z.record(z.string()),
  }), // Outlook triage (move/categorize/flag)
  // An explicit, recorded NON-reversibility marker — the write landed but cannot be undone
  // (resolve-revisions discards the alternative text; send-activity-notification is a fired bell).
  // The advertisement carries reversible:false; this records WHY in the audit trail.
  z.object({ op: z.literal('not-reversible'), reason: z.string() }),
]);
export type InverseDescriptor = z.infer<typeof InverseDescriptorSchema>;

export const ActuationResultSchema = z.object({
  ok: z.boolean(),
  changeId: ChangeIdSchema,
  kind: ActuationKindSchema,
  location: z.string().optional(), // where it landed (range, slide #, comment id, draft id)
  // How to reverse this write (ADR-0007). Populated by the bridge at apply-time once it knows the
  // minted object name / prior state; persisted with provenance so undo is recorded, not implicit.
  inverse: InverseDescriptorSchema.optional(),
  degraded: z.boolean().optional(), // e.g. anchor drifted → applied as a panel item
  // Observability: the reversible write LANDED but its durable provenance could not be persisted
  // (host metadata write failed / unavailable). The change is real but unprovenanced — surface it so
  // the audit trail and the user know, rather than silently dropping the trace. Absent ⇒ recorded.
  provenanceDropped: z.boolean().optional(),
  // Observability: the write LANDED carrying NO provenance payload at all (the turn produced no
  // provenance to stamp). Distinct from `provenanceDropped` (had a record, failed to persist) — this
  // is an unattributed write, surfaced so it is never mistaken for an attributed one.
  provenanceMissing: z.boolean().optional(),
  error: z.object({ code: z.string(), message: z.string() }).optional(),
});
export type ActuationResult = z.infer<typeof ActuationResultSchema>;

/**
 * The read verbs a surface can serve (ADR-0006 capability closure). The CLI grammar scopes host read
 * verbs (`outline`/`read`/`search`) to this list, and conformance tests require a matching bridge
 * read port. `context` is runtime-served and read-only; it may appear here, but it is not a bridge
 * port requirement.
 */
export const ReadVerbSchema = z.enum([
  'outline',
  'read',
  'search',
  'ls',
  'find',
  'tail',
  'list',
  'inspect',
  'properties',
  'comments',
  'attachments',
  'tables',
  'slides',
  'neighbors',
  'context',
  'open',
]);

/** A surface's full capability advertisement: what it can read and what it can write. */
export const CapabilityManifestSchema = z.object({
  surface: SurfaceSchema,
  contextKinds: z.array(ContextKindSchema),
  actuations: z.array(ActuationSchema),
  /** Read verbs this surface serves (ADR-0006); host reads are scoped to this set. */
  reads: z.array(ReadVerbSchema).optional(),
});
export type CapabilityManifest = z.infer<typeof CapabilityManifestSchema>;

/**
 * ADR-0006 — the `Capability` descriptor: the forward source of truth for a single capability.
 *
 * One descriptor names a capability, locates it on a surface, and classifies it on the
 * pure/effect split that ADR-0005 makes load-bearing:
 *   - `read`   — a Layer-B host read (produces a value; never gated).
 *   - `pure`   — a pure transform over values (composes freely; never gated).
 *   - `effect` — an actuation terminal (consumes values, produces a gated `Effect`).
 *
 * The intent (ADR-0006) is that the manifest, the verb→kind map, and dispatch are eventually
 * *derived* from a registry of these descriptors for new capabilities. This is a typed scaffold:
 * no migration is required this wave — the {@link checkCapabilityClosure} conformance gate is what
 * makes that incremental migration safe (drift can't silently return while descriptors and the
 * hand-written manifest/map coexist). `signature` and `gatePolicy` are deliberately open for now;
 * later waves narrow them as the registry lands.
 */
export interface Capability {
  /** Stable capability name (also the CLI/skill identifier, e.g. `reply`, `write-cells`). */
  name: string;
  /** The surface this capability is defined on. */
  surface: Surface;
  /** The pure/effect classification (the ADR-0005 composition + safety boundary). */
  kind: 'read' | 'pure' | 'effect';
  /** A forward type signature for the value layer (open this wave; narrowed later). */
  signature?: unknown;
  /** A forward gate-policy hook for `effect` capabilities (open this wave; narrowed later). */
  gatePolicy?: unknown;
}
