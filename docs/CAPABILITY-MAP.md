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
tested: Word, Excel, Outlook, Teams.** PowerPoint and OneNote remain planned stubs. The remaining gap
to actually load in a host is the **React/Vite/manifest shell** over the web-shell core. 216 tests,
green. See `STATUS.md`.

## Read — attach to context

### Plane A · in-document (Office.js / TeamsJS)
| Capability | Surfaces | Status | Notes |
|---|---|---|---|
| selection, document/body, paragraph | Word | ✅ | `bridge-word` native capture → `@ge/content` blocks |
| table / range / sheet (values+formulas) | Excel | ✅ | `bridge-excel` range/used-range → native table block |
| mail item / subject / from / body | Outlook | ✅ | `bridge-outlook` (string path; reads active item) |
| transcript window | Teams | ✅ | `teams` bridge (RSC-consented, injected) |
| comment / thread | Word, Excel | ✅ | comment-reply actuation + `comment-added` events |
| slide / shape / speaker notes | PowerPoint | ⬜ | `native.slide()` builder ✅ in `@ge/content`; no PPT bridge |
| page / outline | OneNote | ⬜ | modeled; OneNote is web-only; no bridge |
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

## Write — actuate back

### Plane A · in-document
| Actuation (`ActuationKind`) | Surfaces | Status |
|---|---|---|
| **tracked-change** | Word | ✅ `bridge-word` (content-anchored, drift-degrading) |
| **write-cells** | Excel | ✅ `bridge-excel` (address-targeted) |
| **comment-reply** (+ resolve) | Word, Excel | ✅ |
| **reply-mail** | Outlook | ✅ `bridge-outlook` (reviewable `displayReplyForm`) |
| **post-message** (+ Adaptive Card) | Teams | ✅ `teams` (reviewable compose) |
| insert-text, replace-selection, insert-ooxml | Word/PPT | 🟡 modeled |
| fill-content-control | Word | 🟡 |
| insert-slide, set-speaker-notes | PowerPoint | 🟡 modeled; no PPT bridge |
| append-page | OneNote | 🟡 modeled; no bridge |
| create-mail, create-event, create-task | Outlook/Graph | 🟡 modeled |

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
| Provenance write to host durable metadata | 🟡 | `ProvenancePayload` + client `ProvenanceStore` built; per-surface host-metadata write not wired |
| Code-execution file upload (`addContextFile`, v1) | ⬜ | needs a v1 client in `@ge/gemini-client` |
| A2UI render + action→actuation routing | ⬜ | designed (`api/discoveryengine/a2ui.md`); renderer not built |
| Audit | 🟡 | client-direct: provenance-in-artifact is the trail; optional thin sink undecided |

## The gaps that matter (to "map all capabilities" fully)
1. **Sideload shell.** The React/Vite/manifest layer over the `web-shell` core — the last mile to
   load in a real host. (The core logic is built + tested.)
2. **PowerPoint + OneNote bridges.** Real surfaces in the vision; currently planned stubs — the only
   Plane-A reads/writes still 🟡.
3. **Estate writes.** Reads are first-class (`@ge/graph-client`); Graph/SharePoint *writes*
   (send mail, create event, upload/checkout) are modeled but not wired.
4. **Engine paths unbuilt:** v1 `addContextFile` (code-execution uploads) and the A2UI renderer.
5. **Per-capability detection.** Host detection is built; the per-capability intersection gate from
   ADR-0002 is not.

Net: **Word/Excel/Outlook/Teams read + write are live; the estate read/search path is live via
`@ge/graph-client`.** The remaining holes are the sideload shell, the two planned bridges, estate
writes, and the two engine extras.
