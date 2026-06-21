# Capability map — what the add-in can read, write, and do (and what's built)

An honest inventory of the add-in's I/O surface across the **two planes** (in-document via
Office.js/TeamsJS; the estate via Microsoft Graph + Gemini Enterprise connectors), mapped against
what is actually implemented. Companion to `docs/ACCESS-MODEL.md` (the design) and the contract
enums in `@ge/contracts`.

**Status legend**
- ✅ **Built** — implemented + tested in a package.
- 🟡 **Modeled** — typed in `@ge/contracts` and/or request-shapeable, but **not wired to a host**.
- ⬜ **Designed** — described in docs only; no types, no code.

**Where we are:** the *spine* is built (`@ge/contracts`, `@ge/content`, `@ge/gemini-client`). **No
surface bridge exists yet** (`bridge-*`, `teams`, `web-shell` are placeholders), and there is **no
Microsoft Graph client** — so no capability below is yet exercised against a live host. The taxonomy
is largely complete; the wiring is not.

## Read — attach to context

### Plane A · in-document (Office.js / TeamsJS)
| Capability | Surfaces | Status | Notes |
|---|---|---|---|
| selection, document/body, paragraph | Word, all | 🟡 | `ContextKind` modeled; needs a bridge to produce it |
| table / range / sheet (values+formulas) | Excel, Word | 🟡 | native `StructuredData` path in `@ge/content` ✅, but no Excel bridge to feed it |
| slide / shape / speaker notes | PowerPoint | 🟡 | `native.slide()` builder ✅; no PPT bridge |
| comment / thread | Word, Excel, PPT | 🟡 | modeled; no bridge |
| page / outline | OneNote | 🟡 | modeled; OneNote is web-only |
| mail item / thread / attachment | Outlook | 🟡 | modeled; no Outlook bridge |
| calendar event | Outlook | 🟡 | modeled; no bridge |
| transcript window | Teams | 🟡 | modeled; needs RSC consent + bridge |
| image / rendered file | Word/PPT | 🟡 | `file` kind modeled; **note:** rendered-file blob can't attach inline to `streamAssist` (no blob part) |
| prior provenance (read) | Word/PPT/… | ⬜ | not modeled as a read capability yet |

### Plane B · the estate (Microsoft Graph + GE connectors)
| Capability | Source | Status | Notes |
|---|---|---|---|
| **Search SharePoint/OneDrive as the user** | Graph `/search/query` (driveItem, listItem, site) | ⬜ | needs `@ge/graph-client` (delegated NAA token); **not built** |
| Search SharePoint/OneDrive (grounding) | GE **federated** data store | 🟡 | request-shapeable via `dataStoreSpecs`+`filter`; engine does it as the user; not wired to UI |
| Read a specific file (driveItem) | Graph `/drive`, `/sites` | ⬜ | needs `@ge/graph-client`; would resolve to `indexed-document`/`drive-document`/text |
| Read mail / calendar / contacts | Graph `/me/messages`,`/events` | ⬜ | not built (Outlook bridge can read the *active* item via Office.js today) |
| Read the NotebookLM notebook | GE notebook | 🟡 | `UnitDescriptor.notebookId` modeled; engine-side |
| Reference an indexed/Drive doc as context | `documentReference` / `driveDocumentReference` | ✅ | `ContextValue` + `query.parts` mapping built in `@ge/gemini-client` |

## Write — actuate back

### Plane A · in-document
| Actuation (`ActuationKind`) | Surfaces | Status |
|---|---|---|
| insert-text, replace-selection, insert-ooxml | Word/PPT | 🟡 modeled, no bridge |
| **tracked-change** | Word | 🟡 modeled, no bridge |
| fill-content-control | Word | 🟡 |
| comment-reply (+ resolve) | Word/Excel/PPT | 🟡 |
| write-cells, set-entity-card | Excel | 🟡 |
| insert-slide, set-speaker-notes | PowerPoint | 🟡 |
| append-page | OneNote | 🟡 |
| create-mail, reply-mail | Outlook | 🟡 |
| create-event, create-task | Outlook/Graph | 🟡 |
| post-card (Adaptive Card) | Teams | 🟡 |

### Plane B · estate actions (separately authorized)
| Action (`ActionRequest`) | Target | Status | Notes |
|---|---|---|---|
| upload / download / checkout / checkin / add-page | SharePoint, OneDrive | 🟡 | typed in `request.ts` — **but in a separate `ActionRequest`, not unified with `ActuationKind`** (gap) |
| send mail / create event / create task | Graph | ⬜ | not modeled as estate actions; needs `@ge/graph-client` |

## Cross-cutting capabilities
| Capability | Status | Where |
|---|---|---|
| Identity: WIF token exchange (Entra→Google) | ✅ | `@ge/gemini-client` `WifTokenClient` |
| Identity: NAA delegated Entra/Graph token | ⬜ | `AuthClient` not built |
| Grounded streaming (`streamAssist` → SSE events) | ✅ | `@ge/gemini-client` |
| Session continuity / resume | ✅ (client) | `SessionContext` + provenance `sessionId` |
| Connector (data-store) scoping | 🟡 | request shape built; no UI/enumeration |
| Content processing (normalize/chunk/anchor/contextualize) | ✅ | `@ge/content` |
| Context budget + inline/reference/upload strategy | ✅ | `@ge/content` `recommendStrategy`/`ContextBudget` |
| Code-execution file upload (`addContextFile`, v1) | ⬜ | needs a v1 client in `@ge/gemini-client` |
| A2UI render + action→actuation routing | ⬜ | designed (`a2ui.md`); renderer not built |
| Capability **detection** (runtime per-platform) | ⬜ | designed (ADR-0002 / Access Model §6); not built |
| Provenance write to host metadata | 🟡 | `ProvenancePayload` built; `ProvenanceStore` per surface not built |
| Audit | ⬜ | client-direct: provenance-in-artifact is the trail; optional thin sink undecided |

## The gaps that matter (to "map all capabilities" fully)
1. **Estate search/read is not first-class.** Add an `EstateRef` type + a `search` capability to
   `@ge/contracts`, and build **`@ge/graph-client`** (delegated NAA) for `/search/query` + driveItem
   reads — this is what makes "search SharePoint as the user" real on the direct path.
2. **Estate actions are split.** Unify `ActionRequest` (upload/checkout/…) with the `ActuationKind`
   model (or cross-reference them) so write-back is one taxonomy.
3. **No host wiring.** Build the first bridge (Word) + `AuthClient` (NAA) so the 🟡 rows become ✅.
4. **Two engine paths unbuilt:** v1 `addContextFile` (code-execution uploads) and the A2UI renderer.
5. **Capability detection** (the intersection gate from the Access Model) is designed but not coded.

Net: the **taxonomy covers ~all of Plane A and the estate writes**; the **estate read/search path and
all host wiring are the real holes.** The federated-connector search works through the engine with the
request shape we already have; the *direct* Graph search needs `@ge/graph-client`.
