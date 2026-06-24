# Contracts

> **Note (client-direct, `ADR-0001`).** These contracts are the authoritative boundary between the
> **surface-agnostic core** (`runtime`/`web-shell`/`gemini-client`) and the **per-surface bridges**,
> and the shape of every Discovery Engine call. There is **no gateway** — references below to "the
> gateway" or `services/*` endpoints are historical (the add-in calls Discovery Engine directly; see
> the "Gateway endpoints" table caveat at the bottom). The schemas themselves are current: the
> capability grammar, composition expressions, skills, and closure (below) are the live `ADR-0004`→
> `ADR-0006` surface and the single source of truth for the command/expr/skill grammars.

The authoritative boundary between the surface-agnostic core and the per-surface bridges (and the shape of every Gemini Enterprise call). These live in `packages/contracts` as TypeScript types + Zod schemas and are imported by both the core and the bridges. Change them deliberately and update both sides. (Specialist A2A agents, where used, mirror the `Finding` and `UnitDescriptor` shapes.)

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
  slide:   'insert-slide',    // PowerPoint — ADR-0006 CLI parity
  page:    'append-page',     // OneNote   — ADR-0006 CLI parity
  mail:    'reply-mail',      // Outlook   — ADR-0006 CLI parity
  post:    'post-message',    // Teams     — ADR-0006 CLI parity
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

### CLI parity verbs `slide` / `page` / `mail` / `post` (ADR-0006)

These four verbs reach `ActuationKind`s the bridges already **handle and advertise** but that previously had no CLI verb (ADR-0006 closure gaps). Each is gated/approved + provenanced like every other effect (it flows through the Phase-2 plan), is Zod-valid (`ActuationRequestSchema`), and mints its `changeId` exactly once at compile time. All four are **literal-only** this wave (no effect-arg `*Expr`); args are quote-aware via `scanQuoted`.

```
slide "<title>" ["<bullet>" …]   → insert-slide   { params: { slide: { title, bullets[] } } }        (PowerPoint)
page  "<title>" "<body>"          → append-page    { params: { target: { matchText: title }, text: body } }  (OneNote)
mail  "<body>"                    → reply-mail     { params: { mail: { body } } }                       (Outlook)
post  "<text>"                    → post-message   { params: { text } }                                 (Teams)
```

The param shapes match each bridge's `actuate()`/plan: PowerPoint composes a slide from `params.slide`, OneNote takes the page title from `target.matchText` and the body from `params.text`, Outlook builds the draft from `params.mail.body`, Teams stages a reviewable post from `params.text`. (`reply` is already the comment-reply verb, so the Outlook reply verb is `mail`.) Each verb is advertised for a surface only when its mapped kind is in `manifest.actuations[]`, exactly as the other write verbs.

### Named skills: `def` / call (ADR-0005 Phase 3)

A **skill** is a named, parameterized composition the model defines once and calls — the compounding library. This wave delivers **parameterized macros** (substitution only; no `for`/`each` iteration yet).

**Definition** (`packages/contracts/src/skill-grammar.ts`):

```
def <name>(<p1> <p2> …):     ← opens a definition; params are plain or $-prefixed identifiers
  <body line>                ← normal command / composition lines that may reference $p1 … $pN
  …
end                          ← terminates the body
```

`parseProgramBlock` groups a whole `def … end` block into ONE `ParsedSkillDef = { kind: 'skill-def', name, params[], body[] }` (body lines kept verbatim, blanks/comments skipped). A name that **shadows a built-in verb** is rejected (`def set(...)` → corrective); so are a duplicate param, a nested `def`, an unterminated `def` (no `end`), and a stray `end`. AST + Zod: `ParsedSkillDefSchema`, `ParsedSkillCallSchema`.

**Call** — `<name> <arg1> <arg2> …` where `<name>` is a **registered** skill (positional, quote-aware args). `parseProgramBlock` takes the runtime's live skill-name set, so a line whose first token is a registered skill parses as a `ParsedSkillCall` rather than an unknown-verb error; an unregistered name still degrades to the ADR-0004 did-you-mean corrective (back-compat).

**Runtime** (`packages/runtime/src/skill-registry.ts` + `assist-session.ts`):

- A **`SkillRegistry`** holds definitions in an in-session `Map` (durable host-metadata persistence is a tracked follow-up). A `def` line **registers** the skill (no execution → a `skill-registered` confirmation result), validating name/params/body (undeclared `$param` references — other than body-local `let $x` bindings — are rejected at define time).
- A **call** binds `args[i] → params[i]` (exact arity; mismatch / undefined name → corrective), **textually substitutes** each declared `$param` token in the body lines, and re-parses the expanded lines scoped to the registry. The expanded entries run through the **existing Phase-2 plan machinery** (type-check → dry-run → `plan-preview` → one `approvePlan` → gated execute). **A skill call is therefore just a plan: it introduces no new gate, no approval bypass, no effect that skips dry-run.** New `CommandLoopEvent`s: `skill-registered`, `skill-expanded`.

**Safety** (substitution is textual into already-parsed-and-type-checked entries):

- An argument may **not** contain a newline or a code fence (\`\`\`) — so a substituted arg can never inject a new command line or truncate the synthetic fence on re-parse. Only **whole declared-identifier** `$param` tokens are replaced (`$ab` never matches a `$a` binding).
- Every effect among the expanded lines still flows through dry-run + `approvePlan` + the actuation gate; substitution changes *what* a line says, never *whether* it gates. Reject the plan ⇒ the whole expansion is blocked (nothing actuates).
- **Bounded:** the per-turn command budget is decremented for *every* processed entry (including those a call expands into), so an expansion cannot exceed `maxCommandsPerTurn`; the per-plan write cap bounds effects; the skill body is capped (`maxBodyLines`); and skill-call nesting is depth-bounded (`MAX_SKILL_DEPTH`) so a self-/mutually-recursive skill terminates with a corrective rather than recursing unboundedly.

### Read-scoping by `manifest.reads` (ADR-0006)

The grammar advertises a read verb (`outline` / `read` / `search`) **only when it is present in `manifest.reads`** — a surface must never advertise a read it cannot serve. An **absent** `manifest.reads` advertises **no** reads. (Control verbs are always advertised; write verbs remain scoped by their advertised `ActuationKind` as above.)

### Composition: pipelines, bindings, and effect-arg expressions (ADR-0005)

On top of the flat command lines, the grammar has an **expression layer** (`packages/contracts/src/expr-grammar.ts`) — the model can *program* the document. A line is an **expression** (vs. an ADR-0004 simple command) iff it starts with `let ` **or** contains a top-level (unquoted, **unparenthesized**) pipe `|`. Everything else parses exactly as before.

**Phase 1 — pure pipelines + bindings (read-only).** A pipeline is `<source> ( '|' <transform> )*`, where `<source>` is `read <selector>` | `search <text>` | `outline` | `$var`. Reads produce a typed `Value` (`table` | `number` | `text`); pure transforms (`filter`, `select`, `sum`, `avg`, `min`, `max`, `count`, `sort`, `head`, `tail` — `packages/runtime/src/compose.ts`) compose it. `let $name = <pipeline>` binds the resulting `Value` into the loop's env; `$vars` persist across turns **within one `runCommands` loop**. Pipelines are **pure** — they never write. A pipe into an effect (`read X | set …`) is rejected with a corrective; the parser is the structural boundary (`ParsedExpr` = `PipelineExpr | LetExpr`, validated by `ParsedExprSchema`), the runtime evaluator (`evalExpr`) owns transform meaning and the pure-only guard.

**Phase 2 — effect-arg expressions + the gated plan.** An effect verb's value/text arg may be an **expression** evaluated against the binding env, so effects *consume* composed values (the keystone). The rule is deliberately unambiguous (total back-compat):

> An effect arg is an **expression** iff it is exactly **`$var`** or a fully-parenthesized pipeline **`( <pipeline> )`** (an optional leading assignment `= ` is stripped first). Otherwise it is a **literal** — plain text or a `=formula` — exactly as ADR-0004.

Applies to `set` (the value) and `comment` / `reply` (the text); `suggest` / `format` stay literal-only this wave. The parenthesized pipeline is parsed through the same pure pipeline path, so an effect-arg can read+compute but **can never smuggle a write** — a `$var | set …` inside the parens is rejected at eval time, never executed. The parse result carries an optional `valueExpr` / `textExpr: ParsedExpr` beside the verbatim literal:

```ts
// set Summary!B2 = ($anz | sum Revenue)
{ verb: 'set', cell: 'Summary!B2', value: '= ($anz | sum Revenue)',
  valueExpr: { kind: 'pipeline', source: { src: 'var', name: 'anz' },
               stages: [{ name: 'sum', args: 'Revenue' }] } }
// set Sales!F2 =SUM(A1, A2)  → a LITERAL (no valueExpr), unchanged from ADR-0004
{ verb: 'set', cell: 'Sales!F2', value: '=SUM(A1, A2)' }
```

A turn's effects then form a **plan**, executed by `AssistSession.runCommands` (`packages/runtime/src/assist-session.ts`):

1. **Type-check** each effect — its verb's mapped `ActuationKind` must be in `manifest.actuations`; referenced `$vars` must be bound; effect-arg expressions must parse. A failure → a corrective `{error}` for *that* entry; the valid rest still form the plan (never partial execution of a malformed effect).
2. **Dry-run** — execute reads + pure (binding the env), then RESOLVE each effect: evaluate any expression arg via `evalExpr` to a `Value`, render to the concrete param, `compileCommand` → a Zod-validated `ActuationRequest`. **Dry-run actuates nothing.** A non-scalar (`table`) effect value is rejected (a write param is scalar). The resolved set is `PlanEffect[] = { request, command }[]` (`command` is the verbatim line, for the preview).
3. **Plan approval (fail-closed)** — emit a `plan-preview` `CommandLoopEvent`, then call `opts.approvePlan?.(effects)` **once**. No approver ⇒ the whole plan is blocked (each effect → `plan_unapproved`); reject ⇒ all blocked; a thrown approver ⇒ fail closed. On approve, each effect is actuated through the **existing** gate + provenance (the plan approval supersedes the per-write prompt; the trigger gate still runs as the second line). The per-turn write cap is retained; `changeId` is minted once at dry-run and carried unchanged into execution.

Back-compat: when `approvePlan` is absent but `approveWrite` is present, the loop falls back to ADR-0004 per-write approval (Track A). `RunCommandsOptions` gains `approvePlan?: (effects: PlanEffect[]) => boolean | Promise<boolean>`; both `PlanEffect` and the `plan-preview` event are exported from `@ge/runtime`.

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

## Command surface (quick actions · palette · plan)

The typed front of the `/` + `@` pane. All three compile to the same routes — a grounded `send` or
the fail-closed `runCommands` plan gate — so the pane never opens a new actuation path.

### `QuickAction` (`packages/contracts/src/quick-actions.ts`)

A prebuilt button: `{ id, label, surfaces: Surface[], intent: Intent, prompt, ground: string[], output: 'chat' | 'annotation' | 'write', contextMenu: boolean }`. `QUICK_ACTIONS` is the catalog (28); `quickActionsForSurface(surface, allowedIntents?)` filters by surface **and** by capability closure (ADR-0006) so a surface never offers a button it can't honour. `output` decides the route (`chat` → `send`; `write`/`annotation` → the gate); `ground` (e.g. `['this']`, `['unit']`) is prepended as `@`-mentions to form the seed (`quickActionSeed`).

### `CommandPaletteSpec` (`packages/contracts/src/command-palette.ts`)

The `/`-verb list + `@`-mention kinds the input affords: `{ surface, verbs: { intent, label, description }[], mentionKinds: ('document'|'person'|'datastore'|'this'|'upload')[] }`. `commandPaletteFor(surface, allowedIntents?)` returns it, closure-scoped. Each `Intent` maps to a `/label` (`review → /review`, `regen-clause → /rewrite`, `resolve-comment → /resolve`, `draft-slides → /draft`, `meeting-notes → /notes`).

### `CommandPlan` (`packages/contracts/src/command-plan.ts`)

The structured plan the **planner skill** emits, and its parser — a faithful TS port of `skill/m365-command-planner/scripts/parse_plan.py` (kept in lockstep). `CommandPlanSchema = { intent: Intent, surface: Surface, scope?, ground: string[], steps: string[], excludes: string[], clarify: string[], confidence?: 'high'|'medium'|'low' }`. `parsePlanBlock(text) → { plan, errors, needsClarification }`: extracts the ` ```plan ` fence (tolerating an unclosed one), validates intent/surface, accumulates the repeatable keywords, reports unknown keywords with a did-you-mean, and returns a `null` plan on a missing/invalid `intent`/`surface`. A `clarify` line substitutes for a required `step`.

## A2A agent interface (`services/agents`)

Each specialist agent is an ADK agent exposed as an A2A server with an agent card. The gateway calls it as a remote A2A agent (not via StreamAssist `agentsSpec`). Agents accept the resolved unit context + intent and return intent-appropriate output: `review` → `Finding[]`; `resolve-comment` → `{ editedText, replyText, sources }`; `regen-clause` → `{ text, sources }`; `draft-slides` → a stream of slide events; `synthesize` → a stream of tokens + citations. Agents must emit `sources` for every claim; an assertion without a source is a contract violation.

## Gateway endpoints (summary)

> **Historical — no gateway (`ADR-0001`).** In the client-direct architecture the add-in calls
> Discovery Engine `:streamAssist` (and `search`/`completeQuery`/`checkGrounding`/`rank`) **directly**
> as the federated user. The table below maps the *original* gateway routes to the intents they
> served; it survives as a reference for which intent produces which stream shape, not as a live API.
> The optional CORS/audit proxy (`proxyUrl`) is a transparent pass-through, not an intent router.

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
