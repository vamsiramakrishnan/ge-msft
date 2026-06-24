# CAPABILITY-CATALOG.md — the complete cross-surface actuation map

**Status: authoritative.** This is the locked map of every host-native WRITE capability the add-in
models, across Word, Excel, PowerPoint, OneNote, Outlook, and Teams. It is the human-readable face of
`packages/contracts/src/capability.ts` (`ActuationKindSchema` + `ActuationParamsSchema` +
`InverseDescriptorSchema`) — the contract is the typed boundary; this file explains it. Every row was
grounded against the real host typings (`node_modules/@types/office-js/index.d.ts`) or a named
Microsoft Graph endpoint — no capability is asserted without an API behind it.

Read `docs/ADR-0007-host-native-write-kinds.md` first: a write is admitted iff it is **anchored**
(re-resolvable target), **reversible** (a recorded inverse), **provenanced** (agent/sources/identity/
hash in durable metadata), and **gated** (approved before it lands, never auto-applied).

## The two planes

| Plane | What | Auth | Status today |
|---|---|---|---|
| **A — client-direct** | Host writes via Office.js / TeamsJS, in the open document/draft | the signed-in user's Entra/Teams token (already held) | buildable now |
| **B — estate** | Writes to the M365 estate via Microsoft Graph (mailbox, calendar, channels, notebooks server-side) | delegated Graph `*.ReadWrite`/`*.Send` scopes **+ a `GraphClient` write path — neither exists yet** | gated behind the estate-auth path (separate ADR) |

The `graph-client` is **read-only today** (all scopes `*.Read`; `post()` is used only for
`/search/query`). Every Plane-B kind is catalogued but `needs-estate-auth` until a write scope + write
method land. That is a deliberate boundary, not an oversight.

## The closure rule (why unimplemented kinds are safe here)

A kind lives in `ActuationKindSchema` once **modeled**; it is **advertised** by a surface only once
that surface's bridge `actuate()` **handles** it (ADR-0006 capability closure, enforced by the
conformance test comparing each manifest to its `HANDLED_ACTUATIONS`). So the catalog can be complete
while implementation is phased — an unimplemented kind sits in the enum without breaking conformance,
exactly as `create-event`/`post-card` already did. **Locking the map ≠ shipping every kind**; it means
the typed boundary, the inverse strategy, and the verb grammar are decided up front so per-surface
implementation is mechanical and consistent.

Legend — **Status**: `now` = implementable against current typings · `limited` = typings/host-version
constrained (degrade) · `estate` = Plane-B, needs estate-auth · `web-only` = OneNote legacy host.
**Gate**: every write is gated; `on-send` = rides Outlook's existing send veto; `immediate` = commits
on call with no draft buffer; `hard` = irreversible/high-blast-radius, warn explicitly.

---

## Word

Anchor by content (`body.search`) and re-resolve at apply-time; degrade a drifted anchor to a panel
item. Word has **no chart and no equation API** — both are OOXML-only (`insert-ooxml`).

| Capability | Kind | Plane | API (req-set) | Inverse | Gate | Verb | Status |
|---|---|---|---|---|---|---|---|
| Tracked change | `tracked-change` | A | `Range.insertText`/`insertHtml` w/ `trackRevisions` (1.1) | revision reject | yes | `suggest` | ✅ shipped |
| Insert text | `insert-text` | A | `Range.insertText` (1.1) | delete-content | yes | — | ✅ shipped |
| Replace selection | `replace-selection` | A | selection `insertText` replace (1.1) | restore-text | yes | — | ✅ shipped |
| Insert OOXML | `insert-ooxml` | A | `Range.insertOoxml` (1.3) | delete-content | yes | — | ✅ shipped |
| Fill content control | `fill-content-control` | A | `ContentControl.insertText` (1.1) | restore-content-control | yes | — | ✅ shipped |
| Insert rich HTML | `insert-html` | A | `Body.insertHtml` (1.1) / `Range.insertHtml` (1.2) | delete-object(paragraph) | yes | `html` | now |
| Native table | `insert-table` | A | `Body.insertTable` (1.3) | delete-object(word-table) | yes | `wtable` | now |
| Inline image | `insert-image` | A | `Body.insertInlinePictureFromBase64` (1.2) | delete-object(inline-picture) | yes | `image` | now |
| Apply style | `apply-style` | A | `Range.styleBuiltIn` (1.3) / `.style` (1.1) | restore-style | yes | `style` | now |
| Define style | `define-style` | A | `Document.addStyle` (1.5) | delete-object(style) | yes | `define-style` | limited (facets desktop) |
| Numbered/bulleted list | `apply-list` | A | `Paragraph.startNewList`/`List.setLevel*` (1.3) | detach-list | yes | `list` | now |
| Hyperlink | `insert-hyperlink` | A | `Range.hyperlink` (1.4) | restore-text/delete | yes | `link` | now |
| Bookmark | `insert-bookmark` | A | `Range.insertBookmark` (1.4) | delete-object(bookmark) | yes | `bookmark` | now |
| Field (TOC/xref/page#) | `insert-field` | A | `Range.insertField` (1.5) | delete-object(field) | yes | `field` | now |
| New content control | `insert-content-control` | A | `Range.insertContentControl` (1.1/1.5) | delete-object(content-control) | yes | `cc` | now |
| Page/section/line break | `insert-break` | A | `Range.insertBreak` (1.1) | delete-content | yes | `break` | now |
| Page layout / margins | `set-page-layout` | A | `PageSetup.*` (WordApiDesktop 1.2/1.3) | restore-values | yes | `page-layout` | limited (desktop only) |
| Header / footer | `set-header-footer` | A | `Section.getHeader`/`getFooter` (1.1) | restore-text | yes | `header`/`footer` | now |
| Footnote / endnote | `insert-note` | A | `Range.insertFootnote`/`insertEndnote` (1.5) | delete-object(note) | yes | `footnote` | now |
| Accept/reject revisions | `resolve-revisions` | A | `TrackedChange.accept/reject` (1.6); all-scope desktop (1.4) | **not-reversible** | **hard** | `accept`/`reject` | now (per-change) / limited (all) |
| Find & replace | `find-replace` | A | `Body.search` + per-range `insertText` replace (1.1) | restore-text (×N) | yes | `replace-all` | now |
| Document properties | `set-doc-properties` | A | `Document.properties` / custom props (1.3) | restore-doc-properties | yes | `set-property` | now |
| Insert paragraph | `insert-paragraph` | A | `Body.insertParagraph` (1.1) | delete-object(paragraph) | yes | `paragraph` | now |
| Estate doc ops (metadata/export/version) | *(generic Graph)* | B | `PATCH /drives/{id}/items/{id}`, `…/restoreVersion` | restore-values | yes | — | estate |

---

## Excel

Address-anchored (A1/named range), re-resolved at apply-time. The richest client-side write surface —
no tracked-change model, so structural writes are pure wins.

| Capability | Kind | Plane | API (req-set) | Inverse | Gate | Verb | Status |
|---|---|---|---|---|---|---|---|
| Write cells / formulas | `write-cells` | A | `Range.values`/`.formulas` (1.1) | restore-values | yes | `set` / `spill` | ✅ shipped |
| Format cells | `format-cells` | A | `Range.format.*` (1.1) | restore-values | yes | `format` | ✅ shipped |
| Promote to Table | `create-table` | A | `Worksheet.tables.add` (1.1) | delete-object(table) | yes | `table` | Phase B |
| Insert chart | `insert-chart` | A | `Worksheet.charts.add` (1.1) | delete-object(chart) | yes | `chart` | Phase B |
| Conditional format | `format-conditional` | A | `Range.conditionalFormats.add` (1.6) | clear-conditional | yes | `cf` | Phase B |
| Add comment / reply | `add-comment` / `comment-reply` | A | `comments.add`/`reply` (1.10/1.12) | delete-object | yes | `comment`/`reply` | ✅ shipped |
| Linked-entity card | `set-entity-card` | A | linked data types | restore-values | yes | — | limited |
| Estate workbook edits | *(Graph workbook API)* | B | `/workbook/worksheets/{id}/range` | restore-values | yes | — | estate |

---

## PowerPoint

The narrowest namespace: **no** chart-from-data, **no** SmartArt, **no** `addImage` in the namespace,
**no** speaker-notes writer, **no** loose-OOXML insert. The clean wins are shapes, tables, slide
management, and shape formatting — all name/id-anchored and reversible.

| Capability | Kind | Plane | API (req-set) | Inverse | Gate | Verb | Status |
|---|---|---|---|---|---|---|---|
| Insert slide (from base64) | `insert-slide` | A | `slides.add`/`insertSlidesFromBase64` (1.2) | delete-object(slide) | yes | `slide` | ✅ shipped |
| Text box / shape / line | `add-shape` | A | `ShapeCollection.addTextBox`/`addGeometricShape`/`addLine` (1.4) | delete-object(shape) | yes | `shape`/`textbox` | now |
| Native table on slide | `add-table-slide` | A | `ShapeCollection.addTable` (1.8) | delete-object(shape) | yes | `table-slide` | now (1.8) |
| Replace shape text | `set-shape-text` | A | `Shape.textFrame.textRange.text` (1.4) | restore-text | yes | `set-text` | now |
| Format shape | `format-shape` | A | `Shape.fill`/`lineFormat`/`textFrame…font` (1.4) | restore-shape-format | yes | `format-shape` | now |
| Delete slide | `delete-slide` | A | `Slide.delete` (1.2) + `exportAsBase64` (1.8) snapshot | restore-slide | **hard** | `delete-slide` | now (undo 1.8) |
| Reorder slide | `move-slide` | A | `Slide.moveTo` (1.8) | move-slide | yes | `move-slide` | now (1.8) |
| Duplicate slide | `duplicate-slide` | A | `exportAsBase64` + `insertSlidesFromBase64` (1.2/1.8) | delete-object(slide) | yes | `duplicate-slide` | now (1.8) |
| Apply layout | `apply-slide-layout` | A | `Slide.applyLayout` (1.8) | apply-layout | yes | `layout` | now (1.8) |
| Hyperlink | `insert-hyperlink` | A | `TextRange.setHyperlink` (1.10) | delete-object(hyperlink) | yes | `link` | limited (1.10) |
| Image on slide | `insert-image` | A | `setSelectedDataAsync` Image coercion (ImageCoercion 1.1) | *(no handle)* not-reversible | **hard** | `image` | limited (no clean undo) |
| Speaker notes | `set-speaker-notes` | A | **none in typings** | — | — | — | **blocked** (unadvertised) |
| Estate file replace | *(Graph drive item)* | B | `PUT /drive/items/{id}/content` | restore-values (version) | yes | — | estate |

---

## OneNote (web-only, legacy manifest)

`OneNote.run`, web host only; **no event API** (so no `watch`). The weakest reversibility story — most
write methods return objects with **no `.delete()`**; the only in-API deletes are `Paragraph.delete`
and `PageContent.delete`. True page/section delete and targeted replace exist **only on Graph**.

| Capability | Kind | Plane | API (req-set) | Inverse | Gate | Verb | Status |
|---|---|---|---|---|---|---|---|
| Synthesized page | `append-page` | A | `Section.addPage` + `Page.addOutline` (1.1) | *(no Page.delete)* not-reversible | yes | `page` | ✅ shipped |
| Outline on active page | `add-outline` | A | `Page.addOutline` (1.1) | delete-object(page-content) | yes | `outline` | now |
| Append HTML to outline | `append-html` | A | `Outline.appendHtml` (1.1) | *(coarse)* not-reversible | yes | `append` | limited |
| Append paragraph | `append-rich-text` | A | `Outline.appendRichText` (1.1) | delete-object(paragraph) | yes | `append-text` | limited |
| Insert image | `insert-image` | A | `Outline.appendImage` (1.1) | delete-object(paragraph) | yes | `image` | limited |
| Insert table | `insert-table` | A | `Outline.appendTable` (1.1) | delete-object(paragraph) | yes | `table` | limited |
| Create section | `create-section` | A | `Notebook.addSection` (1.1) | *(no delete)* not-reversible | **hard** | `section` | limited |
| Create section group | `create-section-group` | A | `Notebook.addSectionGroup` (1.1) | not-reversible | **hard** | `section-group` | limited |
| Set page title | `set-page-title` | A | `Page.title` (1.1) | restore-text | yes | `title` | now ✦ fully reversible |
| Note tag (To-Do/…) | `add-note-tag` | A | `Paragraph.addNoteTag` (1.1) | delete-object(note-tag) | yes | `tag` | now |
| Insert HTML at cursor | `insert-html-at-cursor` | A | `Application.insertHtmlAtCurrentPosition` | not-reversible | yes | `insert` | limited (no anchor) |
| Embed link at cursor | `insert-link-at-cursor` | A | `Application.insertAndEmbedLinkAtCurrentPosition` | not-reversible | yes | `link` | limited |
| Create page (server) | `graph-create-page` | B | `POST /me/onenote/sections/{id}/pages` | delete-object(onenote-page) | yes | — | estate |
| Targeted patch | `graph-patch-page` | B | `PATCH /me/onenote/pages/{id}/content` | restore-text | yes | — | estate (the reversible OneNote write) |
| Create section (server) | `graph-create-section` | B | `POST /me/onenote/notebooks/{id}/sections` | delete-object(onenote-section) | yes | — | estate |

---

## Outlook

Plane-A writes mutate the **open compose draft** and ride the existing `OnMessageSend` veto gate, so
they are inherently **draft-reviewable**. Importance and follow-up flags are **confirmed absent** from
the compose typings — Graph-only. Graph `sendMail` is the one write that escapes the gate → **not
modeled** (it breaks the no-auto-send invariant).

| Capability | Kind | Plane | API (req-set) | Inverse | Gate | Verb | Status |
|---|---|---|---|---|---|---|---|
| Reply draft | `reply-mail` | A | `item.displayReplyForm` (1.1) | draft-reviewable | on-send | `mail` | ✅ shipped |
| New draft | `create-mail` | A | `mailbox.displayNewMessageForm` (1.6) | draft-reviewable | on-send | `compose` | ✅ shipped |
| Attach file | `add-attachment` | A | `addFileAttachmentFromBase64Async` (1.8) / `addFileAttachmentAsync` (1.1) / `addItemAttachmentAsync` (1.1) | delete-object(attachment) | on-send | `attach` | now |
| Remove attachment | `remove-attachment` | A | `removeAttachmentAsync` (1.1) | draft-reviewable | on-send | `detach` | now |
| Recipients to/cc/bcc | `set-recipients` | A | `to/cc/bcc.setAsync`/`addAsync` (1.1) | restore-values | on-send | `recipients` | now |
| Subject | `set-subject` | A | `subject.setAsync` (1.1) | restore-values | on-send | `subject` | now |
| Body (replace) | `set-body` | A | `body.setAsync` (1.3) | restore-values | on-send | `body` | now |
| Body (prepend) | `prepend-body` | A | `body.prependAsync` (1.1) | restore-values | on-send | `prepend` | now |
| Categories | `add-categories` | A | `item.categories.addAsync` (1.8) | delete-object(category) | **immediate** (read mode) | `categorize` | now |
| Sensitivity label (MIP) | `set-sensitivity-label` | A | `sensitivityLabel.setAsync` (1.13) | restore-values | on-send | `classify` | limited (1.13) |
| Sensitivity class | `set-sensitivity` | A | `sensitivity.setAsync` (1.14) | restore-values | on-send | `sensitivity` | limited (1.14) |
| Custom x-* headers | `set-internet-headers` | A | `internetHeaders.setAsync` (1.8) | delete-object/restore | on-send | `header` | now (x-* only) |
| In-pane notification | `add-notification` | A | `notificationMessages.addAsync` (1.3) | delete-object | **no** (UI affordance) | `notify` | now |
| Appointment compose | `compose-appointment` | A | `displayNewAppointmentForm` (1.1) + start/end/location/attendees setters | draft-reviewable | yes | `meeting` | now |
| Delay delivery | `set-delay-delivery` | A | `delayDeliveryTime.setAsync` (1.13) | restore-values | on-send | `schedule-send` | limited (1.13) |
| Save draft | `save-draft` | A | `item.saveAsync` (1.3) | *(benign)* | no | `save` | now |
| Move message | `move-message` | B | `POST /me/messages/{id}/move` | restore-mail-state | immediate | — | estate |
| Copy message | `copy-message` | B | `POST /me/messages/{id}/copy` | delete-object(message) | immediate | — | estate |
| Create event (server) | `create-event` | B | `POST /me/events` | delete-object(event) | yes | — | estate |
| Categorize/triage received | `categorize-message` | B | `PATCH /me/messages/{id}` | restore-mail-state | immediate | — | estate |
| Flag / follow-up | `flag-message` | B | `PATCH /me/messages/{id}` (flag) | restore-mail-state | immediate | — | estate (no Office.js path) |
| Delete message | `delete-message` | B | `DELETE /me/messages/{id}` (soft) | restore-mail-state | **hard** | — | estate |
| Inbox rule | `create-mail-rule` | B | `POST /me/mailFolders/inbox/messageRules` | delete-object(mail-rule) | **hard** | — | estate (highest privilege) |

> **Excluded by design:** Graph `POST /me/sendMail` (`send-mail`) — the only write that bypasses the
> on-send gate. Modeling it would violate "nothing auto-sends." The reviewable path is
> `create-mail`/`reply-mail` + the on-send gate.

---

## Teams

TeamsJS is mostly **context/UI/navigation** (not content). Every real post/notify/schedule is
**Microsoft Graph (Plane B)** and blocked on the read-only scopes + read-only `GraphClient`. Navigation
primitives (deep links, dialogs, share-to-stage) are *not* gated content writes and are intentionally
**not** in the actuation enum.

| Capability | Kind | Plane | API (scope) | Inverse | Gate | Verb | Status |
|---|---|---|---|---|---|---|---|
| Staged chat post / card | `post-message` | A | TeamsJS `chat.openConversation` / `sharing.shareWebContent` | draft-reviewable | yes | `post` | ✅ shipped |
| Adaptive Card (typed) | `post-card` | A→B | staged (A) / `POST …/messages` w/ card attachment (`ChatMessage.Send`) | draft-reviewable / delete | yes | `card` | A-stage now / estate-send |
| Post chat message | `post-chat-message` | B | `POST /chats/{id}/messages` (`ChatMessage.Send`) | draft-reviewable / soft-delete | yes | — | estate |
| Post channel message | `post-channel-message` | B | `POST /teams/{id}/channels/{id}/messages` (`ChannelMessage.Send`) | soft-delete | yes | — | estate |
| Reply in channel thread | `reply-channel-message` | B | `POST …/messages/{id}/replies` (`ChannelMessage.Send`) | soft-delete | yes | — | estate |
| Update posted message/card | `update-message` | B | `PATCH …/messages/{id}` (channel) | restore-message | yes | — | estate (chat edit API-limited) |
| Reaction | `set-reaction` | B | `POST …/setReaction` (beta) | delete-object(reaction) | yes | — | estate |
| Online meeting | `create-online-meeting` | B | `POST /me/onlineMeetings` (`OnlineMeetings.ReadWrite`) | delete-object(online-meeting) | yes | — | estate |
| Calendar event / invite | `create-event` | B | `POST /me/events` (`Calendars.ReadWrite`) | delete-object(event) | yes | — | estate |
| Task (Planner/To Do) | `create-task` | B | `POST /planner/tasks` / `…/todo/…/tasks` (`Tasks.ReadWrite`) | delete-object(task) | yes | — | estate |
| Activity-feed notify | `send-activity-notification` | B | `POST /users/{id}/teamwork/sendActivityNotification` (`TeamsActivity.Send`) | **not-reversible** | **hard** | — | estate (+ manifest activityType) |

> **Architectural gap (documented, not modeled):** true *interactive* Adaptive Cards (button →
> `Action.Submit` → update-the-card-in-place) require **Bot Framework** server credentials
> (`updateActivity`). That conflicts with the client-direct / no-server-credentials constraint
> (CLAUDE.md), so it is **not** a client-direct kind. Graph can *post* a card but cannot receive its
> action callbacks. Revisit only if an optional bot service is ever added.

---

## Phasing (the build order this map unlocks)

The contract is locked; implementation proceeds per-surface, each kind following the same
`WriteStrategy` pattern (anchor · inverse · provenance · preview) and only advertising once handled.

**Plane A — buildable now (no new auth):**
1. **Phase B — Excel host-native** *(in progress):* `create-table`, `insert-chart`, `format-conditional`.
2. **Phase E1 — Word content:** the 4 shipped text kinds + `insert-table`/`insert-image`/`insert-html`/
   `apply-style`/`apply-list`/`insert-field`/`set-header-footer`/`insert-paragraph` (+ the rest).
3. **Phase E2 — PowerPoint authoring:** `add-shape`/`add-table-slide`/`set-shape-text`/`format-shape`
   + slide management (`delete`/`move`/`duplicate`/`apply-slide-layout`).
4. **Phase E3 — OneNote in-place:** `set-page-title` (fully reversible) → `add-outline`/`insert-table`/
   `insert-image`/`add-note-tag`.
5. **Phase E4 — Outlook compose:** `add-attachment`/`set-recipients`/`set-subject`/`set-body`/
   `add-categories`/`compose-appointment` (all on the on-send gate).

**Plane B — estate (gated behind a new ADR):**
6. **Phase F — estate-write path:** add delegated `*.ReadWrite`/`*.Send` scopes to the manifest, a
   `GraphClient` write method, and the per-call confirm gate (no on-send safety net). Then the Graph
   kinds: Outlook triage/events, Teams posts/meetings/notifications, OneNote server pages/patch.

**Cross-cutting (rides every phase):**
- Extend the **CLI verb grammar** (`packages/contracts/src/command-grammar.ts` `WRITE_VERB_TO_KIND`)
  and the **skill** (`skill/m365-surface-commander`) with each phase's verbs, so streamAssist can emit
  and **compose** them — keep `scripts/parse_commands.py` in lockstep with the TS grammar.
- Thread the new **`InverseDescriptor`** ops through each bridge's apply-time prior-state capture
  (the Word bridge already captures them behind `// TODO(ADR-0007 inverse)` markers).

---

## Counts

The locked catalog spans **~85 actuation kinds**: ~60 Plane-A (client-direct, the bulk buildable now)
and ~25 Plane-B (estate, behind estate-auth). Of Plane A, ~11 are shipped today; the rest are modeled,
typed, and inverse-strategised here — ready for phased, mechanical implementation.
