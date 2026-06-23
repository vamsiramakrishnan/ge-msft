# Contracts

The authoritative boundary between the gateway and the clients. These live in `packages/contracts` as TypeScript types + Zod schemas and are imported by both `services/gateway` and the client packages. Change them deliberately and update both sides. The Python agents in `services/agents` mirror the `Finding` and `UnitDescriptor` shapes.

---

## Intents

```ts
export type Intent =
  | 'assist'           // grounded chat over the unit (StreamAssist)
  | 'review'           // inline review pass → Finding[]  (A2A Review agent)
  | 'resolve-comment'  // edit + reply + resolve a comment
  | 'regen-clause'     // rewrite one content control
  | 'draft-slides'     // generate slides from the unit
  | 'synthesize'       // OneNote page synthesis from the notebook
  | 'meeting-notes';   // Teams live notes + action items
```

## The research unit

```ts
export interface ConnectorRef {
  type: 'sharepoint' | 'onedrive';
  mode: 'federated' | 'ingestion';   // prefer 'federated' for ad-hoc sources
  scope?: string;                    // e.g. "sites/DealRoom"; omit for all the user can see
}

export interface UnitDescriptor {
  notebookId?: string;               // the curated NotebookLM core (precision)
  connectors: ConnectorRef[];        // the live federated edge (breadth)
  restrictToNotebook?: boolean;      // true ⇒ "answer only from the notebook" (regulated work)
  surfaceContext: SurfaceContext;    // what the user is working on
}

export type SurfaceContext =
  | { kind: 'word'; selection?: string; bodyOoxml?: string }
  | { kind: 'excel'; range?: string; values?: string[][] }
  | { kind: 'powerpoint'; slideText?: string }
  | { kind: 'onenote'; pageId?: string; sources?: string[] }
  | { kind: 'teams'; transcriptWindow?: string };
```

## Request

Every grounded call carries the user's Entra/Teams bearer token in `Authorization`; the gateway validates it and federates identity. The body:

```ts
export interface AssistRequest {
  intent: Intent;
  unit: UnitDescriptor;
  query?: string;                    // for 'assist'
  target?: { contentControlId?: string; commentId?: string; range?: string };
  changeId?: string;                 // client-generated; makes write-backs idempotent
}
```

Example (`POST /review`):
```json
{
  "intent": "review",
  "unit": {
    "notebookId": "nb_vendor_risk_7f3",
    "connectors": [{ "type": "sharepoint", "mode": "federated", "scope": "sites/DealRoom" }],
    "restrictToNotebook": false,
    "surfaceContext": { "kind": "word", "bodyOoxml": "<...>" }
  }
}
```

## Findings (agent → gateway → client)

The agent anchors on **content**, never on host range IDs (those are per-session GUIDs the agent can't know). The client resolves to a range with `body.search(matchText)` and re-resolves at apply-time.

```ts
export interface Finding {
  id: string;
  category: 'style' | 'policy' | 'ground';   // ground = verified/positive
  matchText: string;                          // exact text to locate in the host
  contextHint?: string;                       // disambiguates repeated matches
  title: string;
  why: string;
  suggestion?: string;                        // omitted for pure 'ground' findings
  sources: SourceRef[];
  confidence: number;                         // 0..1
  hash: string;                               // for provenance
}

export interface SourceRef { title: string; uri?: string; locator?: string }
```

## SSE event protocol

Streaming endpoints (`/assist`, `/review`, `/draft-slides`, `/synthesize`, `/meeting-notes`) respond with `text/event-stream`. Event types:

```ts
export type SseEvent =
  | { type: 'token'; text: string }                    // incremental text
  | { type: 'finding'; finding: Finding }              // one finding (review)
  | { type: 'slide'; title: string; bullets: string[]; sources: SourceRef[] }
  | { type: 'citation'; source: SourceRef }
  | { type: 'provenance'; payload: ProvenancePayload } // sent before 'done'
  | { type: 'error'; code: string; message: string }
  | { type: 'done' };
```

Wire format: `event: <type>\ndata: <json>\n\n`. Clients consume via `EventSource`; the `StreamClient` falls back to chunked polling of the same payloads if SSE is unavailable.

## Provenance payload

Persisted into the host's durable metadata (Word/PPT custom XML, Excel cell/entity metadata, OneNote page metadata, Teams recap card).

```ts
export interface ProvenancePayload {
  agentId: string;                   // e.g. "review-agent@v2"
  identity: string;                  // signed-in user, e.g. "v.k@acme"
  timestamp: string;                 // ISO 8601
  sources: SourceRef[];
  contentHash: string;               // hash of the generated/edited content
  sessionId?: string;                // StreamAssist session, for resume
}
```

## Write-back actions (separate, explicit)

Connector actions (upload/download/check-in-out) are **never implicit** in a grounding request. They are distinct, explicitly-authorized calls:

```ts
export interface ActionRequest {
  action: 'upload' | 'download' | 'checkout' | 'checkin' | 'add-page';
  connector: 'sharepoint' | 'onedrive';
  target: string;                    // site/path/item
  payload?: { filename: string; contentBase64: string };
  changeId: string;
}
```

## Capability grammar (CLI verbs ↔ actuation kinds)

The client-direct surfaces expose a small, capability-scoped CLI grammar to the model. Each surface advertises only the verbs it can actually serve, derived from its `CapabilityManifest` (`packages/contracts/src/command-grammar.ts`). Three verb classes:

- **Control verbs** (`done`, `help`) — always advertised; not actuations.
- **Read verbs** (`outline`, `read`, `search`) — advertised per `manifest.reads` (see below).
- **Write verbs** — each maps to exactly one `ActuationKind` via `WRITE_VERB_TO_KIND`, and is advertised for a surface ONLY when `manifest.actuations[]` includes its mapped kind. The parser, the grammar advertisement, and the runtime compiler all derive from this single map.

```ts
export const WRITE_VERB_TO_KIND = {
  set:     'write-cells',     // Excel
  suggest: 'tracked-change',  // Word/PPT
  comment: 'add-comment',     // Word/Excel/PPT
  format:  'format-cells',    // Excel
  reply:   'comment-reply',   // Word/Excel — ADR-0006
} satisfies Record<string, ActuationKind>;
```

### `reply <commentId> "text"` → `comment-reply` (ADR-0006)

Replies to an existing comment by its host-opaque id. The first bare token is the comment id (no spaces — host ids like `{3f2a}` or a GUID); the second argument is the quoted reply body (with `\"`/`\\` escapes). It is gated behind the `comment-reply` actuation, which **Word/Excel** advertise. It compiles to:

```ts
// ActuationRequest (validated by ActuationRequestSchema)
{
  changeId,                          // minted exactly once at compile time
  kind: 'comment-reply',
  surface,                           // 'word' | 'excel'
  params: { target: { commentId }, text },
}
```

### Read-scoping by `manifest.reads` (ADR-0006)

The grammar advertises a read verb (`outline` / `read` / `search`) **only when it is present in `manifest.reads`** — a surface must never advertise a read it cannot serve. An **absent** `manifest.reads` advertises **no** reads. (Control verbs are always advertised; write verbs remain scoped by their advertised `ActuationKind` as above.)

### Capability closure

`checkCapabilityClosure({ manifest, handledKinds, readPorts })` (`packages/contracts/src/capability-closure.ts`) is the single, pure definition of whether a surface's *advertised* capability set matches what it can actually *do*. It returns three disagreement sets:

```ts
export interface CapabilityClosureReport {
  phantoms: ActuationKind[];      // advertised actuation kinds the bridge does NOT handle
  unreachedReads: ReadVerb[];     // advertised manifest.reads with no bridge read port
  gaps: ActuationKind[];          // bridge-handled kinds reachable by no CLI write verb
}
```

- **`phantoms` and `unreachedReads` are hard failures** — a surface claiming a write or read it cannot perform is a *lie*; the per-surface conformance test asserts both are empty.
- **`gaps` are tracked, not fatal** — a handler with no CLI verb yet, compared against a checked-in allow-list and burned down deliberately. (`comment-reply` is no longer a gap now that the `reply` verb reaches it.)

`gaps` is computed against `WRITE_VERB_TO_KIND`'s values (the set of kinds some CLI write verb reaches); advertised kinds are de-duped before comparison.

### `Capability` descriptor (forward source of truth)

`packages/contracts/src/capability.ts` introduces a `Capability` descriptor — `{ name, surface, kind: 'read' | 'pure' | 'effect', signature?, gatePolicy? }` — as the eventual single source from which the manifest, the verb→kind map, and dispatch are derived. It is a **typed scaffold only this wave**: no migration is required; the closure conformance gate is what makes incremental migration safe. `signature` and `gatePolicy` are deliberately open (`unknown`) and narrowed in later waves.

## A2A agent interface (`services/agents`)

Each specialist agent is an ADK agent exposed as an A2A server with an agent card. The gateway calls it as a remote A2A agent (not via StreamAssist `agentsSpec`). Agents accept the resolved unit context + intent and return intent-appropriate output: `review` → `Finding[]`; `resolve-comment` → `{ editedText, replyText, sources }`; `regen-clause` → `{ text, sources }`; `draft-slides` → a stream of slide events; `synthesize` → a stream of tokens + citations. Agents must emit `sources` for every claim; an assertion without a source is a contract violation.

## Gateway endpoints (summary)

| Method | Path | Intent | Returns |
|---|---|---|---|
| GET | `/healthz` | — | build info |
| POST | `/assist` | assist | SSE (token, citation, provenance, done) |
| POST | `/review` | review | SSE (finding…, provenance, done) |
| POST | `/resolve-comment` | resolve-comment | `{ editedText, replyText, sources }` |
| POST | `/regen-clause` | regen-clause | SSE (token…, provenance, done) |
| POST | `/draft-slides` | draft-slides | SSE (slide…, provenance, done) |
| POST | `/synthesize` | synthesize | SSE (token, citation…, provenance, done) |
| POST | `/meeting-notes` | meeting-notes | SSE (token/finding…, done) |
| POST | `/action` | — | `{ ok, location }` (connector write-back) |
