# Microsoft 365 add-in capabilities — the platform surface we build on

A reference for what a Microsoft 365 add-in *can* do across Office.js, Outlook, Teams, identity,
and the manifest — mapped to **what this project uses** and **which package owns it**. This is the
"what's possible" companion to `CAPABILITY-MAP.md` ("what we read/write") and `STATUS.md`
("what's built").

Our add-in is an **Office Web Add-in + Teams app** under the **unified Microsoft 365 manifest**
(one package spanning Word, Excel, PowerPoint, Outlook, and Teams), plus a **legacy XML manifest**
for OneNote (web-only). All host access lives ONLY in `bridge-*` / `teams`; the rest of the code is
surface-agnostic.

---

## 1. Extension points — how the add-in shows up

| Extension point | What it is | Hosts | We use it |
|---|---|---|---|
| **Task pane** | An HTML/JS panel docked beside the document; our primary UI (React over the `web-shell` core). | Word, Excel, PowerPoint, OneNote, Outlook, Teams (as a tab) | ✅ primary surface |
| **Function command** | A ribbon button that runs a JS function (no UI) in a shared runtime. | All Office | ◻ for quick actions / commands |
| **Custom functions** | `=NS.FUNC(...)` worksheet functions, incl. **streaming** (`@streaming`). | Excel | ◻ planned (`=GE.ASK` streaming cell) |
| **Event-based activation** | Code that runs on a host *event* without the pane open (a "launch event"). | Outlook (compose/send/new-message), Excel (autorun) | ✅ Outlook on-send gate (`on-send.ts`) |
| **Context menu** | A right-click (shortcut-menu) item that opens the pane or runs a function on the selection. | Word, Excel, PowerPoint (web + Win/Mac M365, perpetual 2021+) | ◻ designed/next |
| **Dialog API** (`displayDialogAsync`) | A modal dialog for auth/consent or rich flows; cross-domain messaging back to the pane. | All Office | ◻ auth fallback when NAA isn't available |
| **Content add-in** | An object embedded *in* the document body. | Excel, PowerPoint | ✕ not used |
| **Contextual / pinnable panes** | Outlook task pane pinned across items; reacts to `ItemChanged`. | Outlook | ✅ `watch()` uses `ItemChanged` |
| **Adaptive Cards / message extension / bot** | Teams surfaces beyond a tab. | Teams | 🟡 `post-message` actuation stages a card; bot/ME not built |

### Context menus (right-click)

A shortcut-menu item gives the selection a fast lane into the add-in. Two action types: **open the
task pane** (`openPage`) or **run a function** with no UI (`executeFunction`). The unified manifest
declares them under `extensions.contextMenus` (a peer of `extensions.ribbons`, schema **v1.23+**);
the legacy XML manifest uses an `ExtensionPoint xsi:type="ContextMenu"` with an `OfficeMenu` target.
Either way the item shares the **same runtime** as the pane — feeding the open pane means going
through the shared runtime and calling `Office.addin.showAsTaskpane()` so the click lands in the
already-mounted `web-shell` rather than a cold load. Hosts are **Word, Excel, PowerPoint** (web +
Win/Mac M365, perpetual 2021+); items are **static** (declared in the manifest, not built at
runtime), only **specific built-in menus** are extendable, and there is **no iPad/mobile** support.
Each item is a `/` command **seeded with the selection as `@this`** — closure-scoped to the surface's
capability set per ADR-0006, so a right-click never offers a verb the bridge can't honor.

## 2. Office.js host object models — read / write / events

Each host exposes a `Host.run(ctx => …)` batch model (Outlook is the exception — callback-based
`Office.context.mailbox.item`). What each surface can do, and what our bridge implements:

### Word — `bridge-word`
- **Read:** selection, body, paragraphs (+ styles → heading levels), content controls, comments,
  `getFileAsync` (full OOXML). → native `Block`s with host locators.
- **Write:** **tracked changes** (`changeTrackingMode`), `insertText`/`insertOoxml`, content-control
  fill, comment replies (+ resolve). We anchor by content (`body.search`), re-resolve at apply-time,
  and degrade to a panel item on drift.
- **Events:** `DocumentSelectionChanged`; `onParagraphChanged/Added/Deleted` (carry a coauthor
  `source`); comments (feature-detected). → `selection-changed` / `document-changed` /
  `comment-added` HostEvents.

### Excel — `bridge-excel`
- **Read:** selected range, used range (values + formulas + address) → native table `Block`.
- **Write:** `write-cells` into an address-targeted range; comment replies.
- **Events:** `worksheets.onSelectionChanged` (address), `worksheets.onChanged` (coauthor
  `EventSource` → origin), `comments.onAdded`.

### PowerPoint — `bridge-powerpoint` *(planned)*
- **Read:** slides, shapes, text, speaker notes. **Write:** `insertSlidesFromBase64`, shape text,
  speaker notes. **Events:** selection changed.

### OneNote — `bridge-onenote` *(planned, web-only)*
- **Read:** page content/outline. **Write:** append page / outline. Ships under the **legacy XML
  manifest** (OneNote isn't in the unified manifest).

### Outlook — `bridge-outlook`
- **Read:** `mailbox.item` — subject, from, body (`getAsync`, HTML/text), recipients; appointments.
  (Read mode is callback-based, not a `run` batch.)
- **Write:** `displayReplyForm` / `displayNewMessageForm` — **reviewable**, never silent send.
- **Events:** `ItemChanged` (pinned pane) → `mail-compose`/`mail-received`; **`OnMessageSend`**
  launch event (Smart Alerts) → the **on-send gate** (`event.completed({allowEvent})`).

### Teams — `teams`
- **Read:** meeting **transcript** (RSC-consented, injected into the bridge), participants.
- **Write:** reviewable **post-message** / Adaptive Card to the chat.
- **Events:** `registerMeetingEndHandler` → `meeting-ended`; session start/end.

## 3. Events we hook (the event-driven layer)

The add-in reacts via `DocBridge.watch(emit)` → `@ge/triggers`. Event sources, by host:

| Source | API | HostEvent | Coauthor-aware |
|---|---|---|---|
| Selection moved | `DocumentSelectionChanged` / `onSelectionChanged` | `selection-changed` | n/a (always local) |
| Document edited | Word `onParagraph*` / Excel `onChanged` | `document-changed` | ✅ `source`/`EventSource` → origin |
| Comment added | Word/Excel comment events | `comment-added` | ✅ |
| Active item switched | Outlook `ItemChanged` | `mail-compose` / `mail-received` | n/a |
| **About to send mail** | Outlook **`OnMessageSend`** (Smart Alerts) | `mail-send` (gate) | n/a |
| Meeting ended | Teams `registerMeetingEndHandler` | `meeting-ended` | n/a |
| Estate changed | Graph change notifications / delta | `estate-changed` | *(planned)* |

**Coauthoring discipline:** content events carry an origin; the registry drops `remote` (a
coauthor's edit or our own write-back) so the system never reacts to itself. The rule lives once in
`coauthorOrigin()` in `@ge/triggers`.

## 4. Identity & data access

| Capability | API | We use it |
|---|---|---|
| **SSO / id token** | MSAL **Nested App Authentication (NAA)** | ✅ `NaaAuthClient.getIdToken()` — the WIF subject |
| Delegated **Microsoft Graph** | NAA access token, delegated scopes | ✅ `getGraphToken(scopes)` → `@ge/graph-client` |
| Interactive fallback | Dialog API / `acquireTokenPopup` | ✅ silent→popup fallback |
| **Federation to Google** | Workforce Identity Federation (RFC 8693 STS), browser-side | ✅ `WifTokenClient` — no Google secret on the client |
| Graph **least privilege** | delegated scopes, **`Sites.Selected`** over all-sites | ✅ scoped in `@ge/graph-client` |
| Graph **search as the user** | `/search/query` (driveItem/listItem/site), `/me/messages`, `/me/events`, `/users` | ✅ `GraphClient` (Plane B) |

Microsoft Graph is reached **directly from the add-in** as the signed-in user (delegated) — not via
Gemini Enterprise's connectors. (Gemini's *federated data stores* are a separate, engine-side path
for grounding, scoped per request with `dataStoreSpecs`.)

## 5. Teams app capabilities

| Capability | What it enables | Status |
|---|---|---|
| **Tab** (personal/channel) | Host the task-pane app inside Teams | ✅ (web-shell renders) |
| **Meeting app** (side panel / in-meeting / stage) | Live meeting context + actions | 🟡 transcript path built |
| **Bot** (Bot Framework) | Conversational + proactive messages | ⬜ not built |
| **Message extension** | Compose/search actions from the compose box | ⬜ not built |
| **RSC** (Resource-Specific Consent) | Read the meeting **transcript** without org-wide admin grants | ✅ assumed for transcript capture |
| **TeamsJS** | Teams client SDK (context, meeting events) | ✅ feature-detected `TeamsJsLike` |

## 6. Manifest & packaging

- **Unified Microsoft 365 manifest** (`manifests/m365-unified.manifest.json`) — one package spanning
  Word, Excel, PowerPoint, Outlook, **and** Teams (tab/meeting/bot/ME). Declares task panes, function
  commands, **event-based activation** (Outlook launch events), requirement sets, and Graph
  permissions.
- **Legacy XML manifest** (`manifests/onenote.manifest.xml`) — OneNote ships separately (web-only,
  not in the unified manifest).
- **Requirement sets** gate which APIs exist on a given host/build — we **feature-detect** at runtime
  (e.g. `watch()` skips events an older set lacks) and `detectSurface()` picks the bridge.

## 7. Constraints that shape our design

- **No inline blob to `streamAssist`** (`v1alpha` `query.parts` has no base64/media part) → the
  **reference-over-inline** policy; large/binary objects attach as an indexed reference or extracted
  text, not raw OOXML.
- **Code-execution file uploads** are a separate **v1** `addContextFile` path (Python, ~30s runtime)
  — designed, not built.
- **Outlook on-send** must **fail open** if *deciding* throws (never wedge Send), but a genuine
  **block must not be downgraded** to allow — see `on-send.ts`. A bounded timeout for a *hung*
  trigger is a noted follow-up.
- **OneNote is web-only**; **Word body events** vs **Excel `EventSource`** differ in coauthor
  signaling; **Outlook reads are callback-based** — all reasons surface specifics stay in the bridges.
- **Treat all host/document/transcript content as untrusted data**, never instructions (screened by
  Model Armor at the engine; framed as data in `query.parts`).

---

### Capability → API → package → status (quick index)

| Capability | Microsoft API | Package | Status |
|---|---|---|---|
| Read selection/doc/paragraphs | Word.run object model | `bridge-word` | ✅ |
| Read range/sheet | Excel.run object model | `bridge-excel` | ✅ |
| Read mail item | `mailbox.item` getAsync | `bridge-outlook` | ✅ |
| Read transcript | TeamsJS (RSC) | `teams` | ✅ |
| Read slides/notes | PowerPoint.run | `bridge-powerpoint` | ⬜ planned |
| Read OneNote page | OneNote.run | `bridge-onenote` | ⬜ planned |
| Search/read estate as user | Graph `/search`, `/drive`, `/messages`, `/events` | `graph-client` | ✅ |
| Tracked-change write | Word.run + changeTracking | `bridge-word` | ✅ |
| Write cells | Excel.run range | `bridge-excel` | ✅ |
| Reviewable reply | `displayReplyForm` | `bridge-outlook` | ✅ |
| Post to chat / card | TeamsJS / Adaptive Card | `teams` | ✅ |
| Selection/edit/comment events | host change events | `bridge-*` `watch()` | ✅ (Word/Excel/Outlook/Teams) |
| On-send gate | `OnMessageSend` launch event | `bridge-outlook` `on-send.ts` | ✅ |
| SSO id token | MSAL NAA | `web-shell` `NaaAuthClient` | ✅ |
| Delegated Graph token | MSAL NAA | `web-shell` + `graph-client` | ✅ |
| Federate Entra→Google | WIF STS (RFC 8693) | `gemini-client` `WifTokenClient` | ✅ |
| Custom function `=GE.ASK` | Excel custom functions (streaming) | `bridge-excel` | ⬜ planned |
| Bot / message extension | Bot Framework / TeamsJS | `teams` | ⬜ planned |
