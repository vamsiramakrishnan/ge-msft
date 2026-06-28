# Capability map — what the add-in can read, write, and do (and what's built)

An honest inventory of the add-in's I/O surface across the **two planes** (in-document via
Office.js/TeamsJS; the estate via Microsoft Graph + Gemini Enterprise connectors), mapped against
what is actually implemented. Companion to `ADR-0002-capability-model.md` (the design),
`MICROSOFT-ADDIN-CAPABILITIES.md` (the platform surface), and the contract enums in `@ge/contracts`.

**Status legend**
- ✅ **Built** — implemented + tested in a package.
- 🟡 **Modeled** — typed in `@ge/contracts` and/or request-shapeable, but **not yet wired to a host**.
- ⬜ **Designed** — described in docs only; no types, no code.

**Where we are:** the spine **and the first wave of host wiring** are built. Surface-agnostic core:
`@ge/contracts`, `@ge/content`, `@ge/gemini-client` (streamAssist + search/autocomplete/grounding/
rank), `@ge/graph-client` (delegated Graph reads), `@ge/triggers` (event engine + actuation gate),
`@ge/runtime` (AssistSession + ContextModel + Orchestrator), and the `@ge/web-shell` core
(`NaaAuthClient`, `composeSession`, `PanelController`, `ProvenanceStore`). **Surface bridges built and
tested: Word, Excel, Outlook, Teams, PowerPoint, OneNote** — each with an advertised==handled
capability set, conformance-enforced per ADR-0006. The remaining gap to actually load in a host is
the **React/Vite/manifest shell** over the web-shell core. See `STATUS.md`.

## Read — attach to context

### Plane A · in-document (Office.js / TeamsJS)
| Capability | Surfaces | Status | Notes |
|---|---|---|---|
| selection, document/body, paragraph | Word | ✅ | `bridge-word` native capture → `@ge/content` blocks |
| table / range / sheet / named range (values+formulas) | Excel | ✅ | `bridge-excel` range/used-range + workbook tables/names → native table block and openable refs |
| mail item / subject / from / body | Outlook | ✅ | `bridge-outlook` (string path; reads active item) |
| transcript window | Teams | ✅ | `teams` bridge (RSC-consented, injected) |
| comment / thread | Word, Excel | ✅ | comment-reply actuation + `comment-added` events |
| slide / shape | PowerPoint | ✅ | `bridge-powerpoint` (selected slides → shapes' text) |
| page / outline | OneNote | ✅ | `bridge-onenote` (active page title + outline rich text; web-only) |
| calendar event (active item) | Outlook | 🟡 | Office.js appointment read not yet in the bridge |
| image / rendered file | Word/PPT | 🟡 | `file` kind modeled; **note:** can't attach inline to `streamAssist` (no blob part) |
| prior provenance (read) | Word/PPT/… | 🟡 | `ProvenanceStore` (client view-model) built; durable host-metadata read per surface not wired |

### Plane B · the estate (Microsoft Graph + GE connectors)
| Capability | Source | Status | Notes |
|---|---|---|---|
| **Search SharePoint/OneDrive as the user** | Graph `/search/query` (driveItem, listItem, site) | ✅ | `@ge/graph-client` (`search` → `EstateRef[]`), delegated NAA token |
| Read a specific file (driveItem) | Graph `/drive`, `/sites` | ✅ | `getDriveItem` → `driveItemToContext` |
| Read mail / calendar / people | Graph `/me/messages`, `/me/events`, `/users` | ✅ | `getMessage`/`getEvent`/`getUser` → context mappers |
| Search SharePoint/OneDrive (grounding) | GE **federated** data store | 🟡 | request-shapeable via `dataStoreSpecs`+`filter`; engine does it as the user; not wired to UI |
| Read the NotebookLM notebook | GE notebook | 🟡 | `UnitDescriptor.notebookId` modeled; engine-side |
| Reference an indexed/Drive doc as context | `documentReference` / `driveDocumentReference` | ✅ | `ContextValue` + `query.parts` mapping in `@ge/gemini-client`; search hits map straight to references |

### Read CLI verbs (`outline` / `read` / `search`) — `CapabilityManifest.reads` per surface (ADR-0006)
These are the addressable read verbs the command grammar advertises per surface — distinct from
context-attach (`listContext`/`resolveContext`), which is the **universal** port every bridge serves
and is NOT a read verb. ADR-0006 closure: a surface advertises a read verb ONLY when it has the
matching bridge port, and a per-surface conformance test (`capability-closure.test.ts`) fails the
build on drift.

| Surface | `reads` | Ports backing them |
|---|---|---|
| Word | `outline`, `read`, `search` | `captureDocState` (outline + whole-doc `read`), `searchDocument` |
| Excel | `outline`, `read`, `search` | `captureDocState` (outline), `readRange` (addressable `read <A1\|NamedRange>`, bounded to `MAX_READ_CELLS`=10k cells), `searchDocument` |
| PowerPoint | `outline`, `read`, `search` | `captureDocState` (slide inventory, bounded `MAX_READ_SLIDES`=60), `readRange` (addressable `read <slide:N>`, single slide), `searchDocument` (slide-text scan, `MAX_SEARCH_SLIDES`=8) |
| OneNote | `outline`, `read`, `search` | `captureDocState` (active-page title + paragraph outline; also backs whole-page `read` — no addressable sub-range), `searchDocument` (page-paragraph scan, `MAX_SEARCH_PARAGRAPHS`=8) |
| Outlook | `read`, `search` | `captureDocState` (whole-item `read`: subject + from + leading body lines, bounded `MAX_OUTLINE_LINES`=40 — a mail item has no sub-range), `searchDocument` (body-line scan, `MAX_SEARCH_LINES`=8). No `outline` (no heading structure). |
| Teams | `read`, `search` | `captureDocState` (whole-transcript `read`: meeting title + turn lines, bounded `MAX_TRANSCRIPT_LINES`=60 — a transcript has no sub-range), `searchDocument` (transcript-line scan, `MAX_SEARCH_LINES`=8). No `outline` (no heading structure). |

All four surfaces' ports are **read-only, bounded, and item/window-scoped** (Outlook reads only the
active `mailbox.item`; Teams only the in-memory transcript window; PowerPoint/OneNote only the active
deck/page), degrade to `[]`/`undefined` on bad input (older host via `isSet(...)`, unaddressable
selector, no active item), and frame host content strictly as data (native/string `@ge/content`
path or `buildDocStateSnapshot`'s untrusted-wrapped envelope). A surface advertises a read verb ONLY
when its bridge method exists — conformance-enforced per surface (`capability-closure.test.ts`).

## Write — actuate back

### Plane A · in-document
| Actuation (`ActuationKind`) | Surfaces | Status |
|---|---|---|
| **insert-text** | Word | ✅ `bridge-word` direct plain-text insert, current selection or exact anchor; irreversible until inserted-range inverse exists |
| **replace-selection** | Word | ✅ `bridge-word` direct replacement of the current selection; prior text captured as inverse |
| **insert-ooxml** | Word | ✅ `bridge-word` direct rich OOXML insert, current selection or exact anchor; irreversible until inserted-range inverse exists |
| **tracked-change** | Word | ✅ `bridge-word` (content-anchored, drift-degrading) |
| **add-comment** | Word, Excel | ✅ content/cell-anchored new comment (ADR-0004 `comment` verb) |
| **write-cells** | Excel | ✅ `bridge-excel` (address-targeted) |
| **format-cells** | Excel | ✅ `bridge-excel` (ADR-0004 `format` verb) |
| **create-table** | Excel | ✅ `bridge-excel` native table creation |
| **insert-chart** | Excel | ✅ `bridge-excel` native chart creation over a source range |
| **format-conditional** | Excel | ✅ `bridge-excel` conditional-format rules |
| **comment-reply** (+ resolve) | Word, Excel | ✅ |
| **fill-content-control** | Word | ✅ `bridge-word` direct fill of a known content-control id; prior value captured as inverse |
| **insert-slide** | PowerPoint | ✅ `bridge-powerpoint` (native compose or client-staged base64 PPTX deck import) |
| **set-shape-text** | PowerPoint | ✅ `bridge-powerpoint` exact slide+shape text replacement |
| **append-page** | OneNote | ✅ `bridge-onenote` (synthesis + inline citation tags) |
| **reply-mail** | Outlook | ✅ `bridge-outlook` (reviewable `displayReplyForm`) |
| **create-mail** | Outlook | ✅ `bridge-outlook` (reviewable `displayNewMessageForm`; `compose` verb, unaddressed by default) |
| **post-message** (+ Adaptive Card) | Teams | ✅ `teams` (reviewable compose) |
| set-speaker-notes | PowerPoint | 🟡 modeled — **not advertised** (no host write path in current typings; always degraded, so un-advertised per ADR-0006) |
| add-shape, add-table-slide, format-shape, delete/move/duplicate slide | PowerPoint | 🟡 modeled — **not advertised** until the bridge has tested host write paths and inverses |
| create-event, create-task | Outlook/Graph | 🟡 modeled — **not advertised** (no `actuate()` case yet) |

**Every write verb maps to a CLI verb and is composition-bearing.** `WRITE_VERB_TO_KIND`:
`set`→write-cells, `suggest`→tracked-change, `comment`→add-comment, `format`→format-cells,
`reply`→comment-reply, `table`→create-table, `chart`→insert-chart, `cf`→format-conditional,
`spill`→write-cells, `slide`→insert-slide, `shape`→set-shape-text, `page`→append-page,
`mail`→reply-mail, `post`→post-message, `compose`→create-mail. Bridge-backed non-core actuation
kinds, such as Word `insert-text` and `fill-content-control`, are exposed as exact slash commands
(`/insert-text`, `/fill-content-control`) only on surfaces that advertise them. The free-text slot of
every text-bearing verb accepts a
`( <pipeline> )` / `$var` expression evaluated at dry-run (ADR-0005), so a composed value can feed a
cell, a comment, a slide's bullets (table rows → bullets), a page, an email, or a chat post — not
just a cell. Only `suggest`/`format` (no free-text slot) stay literal.

### Plane B · estate actions (separately authorized)
| Action | Target | Status | Notes |
|---|---|---|---|
| upload / download / checkout / checkin / add-page | SharePoint, OneDrive | 🟡 | typed in `request.ts`; estate **reads** are built, estate **writes** are not wired |
| send mail / create event / create task | Graph | 🟡 | reads built (`graph-client`); writes not yet |

## Cross-cutting capabilities
| Capability | Status | Where |
|---|---|---|
| Identity: WIF token exchange (Entra→Google) | ✅ | `@ge/gemini-client` `WifTokenClient` (epoch-safe invalidation) |
| Identity: NAA delegated Entra id + Graph token | ✅ | `@ge/web-shell` `NaaAuthClient` |
| Grounded streaming (`streamAssist` → SSE events) | ✅ | `@ge/gemini-client` |
| **Faceted/filtered search, autocomplete, rerank, grounding-check** | ✅ | `@ge/gemini-client` (`search`/`completeQuery`/`rank`/`checkGrounding`) |
| Session continuity / resume (cross-surface) | ✅ | `AssistSession` + `resumeSessionId` + provenance `sessionId` |
| **Event-driven reactions** (selection/edit/comment/on-send/meeting) | ✅ | `bridge-*` `watch()` → `@ge/triggers` → `@ge/runtime` Orchestrator |
| **Working-context construction → session commit** (fold/prime) | ✅ | `@ge/runtime` `ContextModel` + `AssistSession.commit` |
| **Actuation gate** (on-send / pre-actuation veto) | ✅ | `@ge/triggers` `TriggerRegistry.gate` + Outlook `on-send.ts` |
| Content processing (normalize/chunk/anchor/contextualize) | ✅ | `@ge/content` |
| Context budget + inline/reference/upload strategy | ✅ | `@ge/content` `recommendStrategy`/`ContextBudget` |
| Connector (data-store) scoping | 🟡 | request shape built; no UI enumeration |
| Capability **detection** (runtime) | 🟡 | `detectSurface()` (host→bridge) built; per-capability intersection gate not |
| Provenance write to host durable metadata | 🟡 | Wired for **Word** (custom XML part) + **Excel** (workbook settings) via `provenance-record.ts`; **not yet** for PPT/OneNote/Outlook/Teams. A landed-but-untraced write is now **observable**: `ActuationResult.provenanceDropped` (had a record, failed to persist) and `provenanceMissing` (no payload at all → unattributed) are surfaced on the run-step instead of swallowed. Client `ProvenanceStore` view-model built for all. |
| Code-execution file upload (`addContextFile`, v1) | ✅ | `@ge/gemini-client` `ContextFileClient` + runtime `context` strategy; host/user attachment still explicit |
| A2UI render + action→actuation routing | ⬜ | designed (`api/discoveryengine/a2ui.md`); renderer not built |
| Audit | 🟡 | client-direct: provenance-in-artifact is the trail; optional thin sink undecided |

## The gaps that matter (to "map all capabilities" fully)
1. **Sideload shell.** The React/Vite/manifest layer over the `web-shell` core — the last mile to
   load in a real host. (The core logic is built + tested.)
2. **CLI parity — closed.** Every handled actuation now has a model-facing CLI verb
   (`slide`/`page`/`mail`/`post`/`compose`), and every text-bearing effect verb is
   composition-bearing (ADR-0005 expressions in the free-text slot). The remaining write frontier is
   not *more verbs* but **deeper host operations** per surface (Word styles/headers/revisions,
   richer PowerPoint object creation/formatting, Outlook calendar/tasks) and **cross-surface composition** (a multi-bridge router so a
   value read on one surface can be written on another — today an `AssistSession` owns one bridge;
   live cross-surface needs Graph/the research unit because each add-in instance runs in one host).
3. **Estate writes.** Reads are first-class (`@ge/graph-client`); Graph/SharePoint *writes*
   (send mail, create event, upload/checkout) are modeled but not wired.
4. **Engine paths unbuilt:** the A2UI renderer. `addContextFile` client support exists; live host UX
   still needs the explicit attach/upload affordance.
5. **Per-capability detection.** Host detection is built; the per-capability intersection gate from
   ADR-0002 is not.

Net: **All six surface bridges (Word/Excel/PowerPoint/OneNote/Outlook/Teams) have a truthful,
conformance-enforced read+write set, every write verb is CLI-reachable and composition-bearing, and
the estate read/search path is live via `@ge/graph-client`.** The remaining holes are the sideload
shell + live-host validation, cross-surface composition, deeper per-surface host operations, estate
writes, and the two engine extras.
