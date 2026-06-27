# Mechanisms to provide context (and scope connectors)

How the client-direct add-in hands context to Gemini Enterprise and scopes which sources ground a
turn. Everything here is a **field on the next `streamAssist` request** — no backend, no server
state. Grounded in the `v1alpha` schema (see `streamAssist.md`).

## 1. `query.parts[]` — the attach mechanism

A `streamAssist` request's `query` may carry `parts[]`, each one of:

| Part | Fields (input) | Use for |
|---|---|---|
| **text** | `text`, `mimeType` | Selections, ranges, transcript windows, email bodies — extracted host content as **data**. |
| **documentReference** | `documentName` (VAIS doc), `displayTitle`, `urlForConnector` | An item already **indexed** in a connected data store — preferred over inline text (ACL-preserving, citations resolve). |
| **driveDocumentReference** | `driveId`, `documentName`, `displayTitle` | A Google Drive document. |
| **personReference** | `displayName`, `email`, `personId` | A person/contact reference. |
| **uiJsonPayload** | `text` (JSON) | A2UI user-interaction messages only (see `a2ui.md`). |

There is **no inline blob/base64 part and no REST media-upload for session files in `v1alpha`** —
so binary/large objects (PDF/PNG/OOXML) are attached either as extracted **text** or, preferably, as
a **reference** to their indexed copy. This is the load-bearing constraint behind the
*reference-over-inline* policy.

### Reference-over-inline (the resolution policy)

When a piece of content — active document (Plane A) **or** estate (Plane B) — already exists as an
indexed document in a connected data store, attach it as a `documentReference`, **not** as extracted
text. Only fall back to a text part for ad-hoc items no connector covers. Benefits: preserves ACL
grounding, keeps citations resolvable, and stays within the parts budget (a 40-page contract open in
Word grounds better by referencing its indexed SharePoint copy than by inlining its OOXML).

### Metadata to add per part

`mimeType` (`text/markdown`, `text/csv`, `text/html`), a human `displayTitle`, source identity
(`documentName`/`urlForConnector`), and structural hints (heading path, sheet+range header, slide
index, sender/recipients/timestamps). Chunk long bodies into labeled section-parts so citations are
granular. See the "Parsing content + metadata" notes in the design discussion / `ACCESS-MODEL`.

## 2. `toolsSpec.vertexAiSearchSpec` — connector (data-store) scoping

"Connectors" are Discovery Engine **data stores**. Scope them **per request** — toggling a connector
chip adds/removes a `dataStoreSpec`:

```jsonc
"toolsSpec": {
  "vertexAiSearchSpec": {
    "dataStoreSpecs": [
      { "dataStore": "projects/…/collections/…/dataStores/sharepoint_dealroom",
        "filter": "site = \"DealRoom\"", "boostSpec": { /* prefer, don't exclude */ } },
      { "dataStore": "projects/…/dataStores/vendor_policies" }
    ],
    "filter": "notebookId: ANY(\"nb_vendor_risk_7f3\")"   // restrictToNotebook
  }
}
```

- **All sources the user can see** → omit `dataStoreSpecs` (engine default).
- **Specific selection** → list the checked data stores; unchecking one = leaving it out.
- **Per-store `filter`** narrows within a store; the top-level `filter` scopes the whole turn (e.g.
  "answer only from the notebook").
- **`boostSpec`** lets a toggle *prefer* a source rather than hard-excluding the rest.

Enumerate the available stores once (`engines.get` / `dataStores.list`) to draw the chips; the
selection lives in the request builder, not on a server.

## 3. `actionSpec` — enable/disable agent actions

`StreamAssistRequest.actionSpec.actionDisabled: true` turns off action-taking for the turn
(enterprise edition). This is the natural backing for a global **"read-only vs can-act"** switch in
the UI.

## 4. Sessions — multi-turn + cross-surface resume

Pass `session` to carry conversation state; read `sessionInfo.session` from the response and persist
it into host metadata (provenance `sessionId`) so a reopened artifact — or the same unit opened on
another surface — resumes. See `sessions.md`.

## 5. `generationSpec.modelId` — model override

Optional per-turn model selection; otherwise the engine's configured default is used.

## Summary

Attaching context = `query.parts[]` (prefer references). Scoping connectors = `dataStoreSpecs` +
`filter` + `boost`. Enabling agents to act = `actionSpec`. Continuity = `session`. All request
fields; the add-in is a request builder + a stream renderer.
