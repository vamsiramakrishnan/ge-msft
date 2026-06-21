# ADR-0002 — The capability model (context capture + actuation)

**Status:** Accepted (2026-06-21) · builds on ADR-0001.

## Context

The add-in's value is not "a chat box in five apps." It's that each Microsoft surface can
**lift its rich host objects into a live Gemini Enterprise session as context**, and **actuate**
the agent's output back into the host — reversibly and with provenance. Experiences and chosen
agents (via connectors / custom agents configured in the engine) are then composed *on top of*
these foundational capabilities, not hard-coded.

This is a deliberate inversion: build the **capabilities** (read objects → context, write objects
← actuation) as the stable foundation; treat workflows as configuration over them.

## What the API actually allows (grounding the design)

From the Discovery Engine `v1alpha` surface (see `docs/api/discoveryengine/`):

- **Live context attaches via `query.parts[]`** on `streamAssist`. Each part is one of:
  `text` (with `mimeType`), `documentReference` (a doc already indexed in a connected data store),
  `driveDocumentReference`, or `personReference`. There is **no public REST media-upload** for
  session files in `v1alpha`, so binary/large objects are attached either as extracted text or as
  references to already-indexed documents.
- **Sessions** carry the context across turns; `sessionInfo.session` is persisted into host
  metadata (provenance) so a reopened artifact resumes.

So the foundation maps cleanly: a surface's "attachable objects" become parts; the engine grounds
on them as the signed-in user.

## Decision

Two foundational, surface-agnostic capability surfaces, defined in `@ge/contracts`:

1. **Context capture** (`context.ts`) — `Surface`, `ContextKind` (selection, document, table,
   range, slide, mail-item, mail-thread, attachment, calendar-event, transcript, person,
   indexed-document, drive-document, …), `ContextRef` (a cheap UI handle), and `ResolvedContext`
   whose `value` maps 1:1 to a `query.parts[]` entry.
2. **Actuation** (`capability.ts`) — `ActuationKind` (tracked-change, insert-ooxml,
   fill-content-control, write-cells, insert-slide, set-speaker-notes, reply-mail, create-event,
   create-task, post-card, …), `ActuationRequest` (idempotent `changeId` + optional
   `ProvenancePayload`), `ActuationResult`, and `CapabilityManifest` (what a surface advertises it
   can read and write).

The engine that turns attached context into a session lives in `@ge/gemini-client`:
**`SessionContext`** holds the live attached set and emits `query.parts[]`; `streamAssist` accepts
a `context` option and sends a multi-part query (each object as data, the user's question last).

Each surface bridge (`packages/bridge-*`, `packages/teams`) implements one `DocBridge`:
`getCapabilities()`, `listContext()` / `watchContext()`, `resolveContext(ref)`, and
`actuate(request)`. Bridges are the *only* code that touches Office.js/TeamsJS/Graph.

## UX consequence

The shared panel is organized around these primitives, not around a single workflow:

- A **context tray** — chips for what's attached now, with live re-resolution (e.g. "current
  selection"), add/remove, and a budget indicator. This is the "add objects to the session" UI.
- An **actuation review** — the agent proposes writes; the user sees a reversible, provenanced
  diff/card before it lands. This is the "actuate through the add-in" UI.
- A **command palette** of the surface's advertised actuations, identical in shape across apps.

## Consequences

- Adding a surface = implementing one `DocBridge` + declaring its `CapabilityManifest`. No gateway,
  no new backend.
- Adding an experience = composing existing context kinds + actuations (+ an engine-side agent),
  not new plumbing.
- Provenance and reversibility are enforced at the actuation boundary, uniformly.
