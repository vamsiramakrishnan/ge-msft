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
