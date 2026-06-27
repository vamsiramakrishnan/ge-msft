# ADR-0003 — Document-as-environment context construction

**Status:** Accepted (2026-06-22) · refines the anchoring/grounding contract in `01-architecture.md` and `CONTRACTS.md`; complements ADR-0001 (client-direct) and ADR-0002 (capability model).

## Context

A teardown of Anthropic's shipped **Claude for Microsoft 365** add-in (the canonical manifest at
`pivot.claude.ai` and its public client bundle) showed a context-construction approach materially
different from ours, and worth learning from:

1. **The document is an addressable environment, not a payload.** Their client mounts the open
   document read-only into an in-browser virtual filesystem (sheets as CSV, the doc as an outline,
   the deck as slides) and lets the model read it **lazily** with ordinary file tools
   (`ls`/`wc`/`head`/`cat`) plus a Python sandbox, rather than pre-serialising the whole file into
   the prompt. On a 50–70k-row workbook this is dramatically more token-efficient.
2. **An ambient `<doc_state>` snapshot** carries structure every turn — sheet/slide inventory,
   named ranges, the current selection, and comments — so the model always knows the shape of the
   document without reading it.
3. **Writes are auditable by construction.** Their prompts enforce "default to spreadsheet
   formulas" (a result the user can inspect, not an opaque literal), **tracked changes**, and a
   **source comment on every externally-sourced cell** — citations rendered as host comments.
4. **Long sessions are compacted**, and host content is treated as **untrusted** (data, never
   instructions).

Our current pipeline (`@ge/content`: native object model → blocks → **budgeted chunks** → grounded
`streamAssist`) is solid for grounding on the **unit** (notebook + federated SharePoint/OneDrive),
which is our structural advantage — Claude is *active-file-only*. But for the **working document**
it pre-chunks eagerly, has no ambient structural snapshot, no lazy on-demand reads, and no
formula-first / comment-as-citation write discipline. That is exactly the depth Claude's design
buys on large, computational documents.

We are also not Anthropic's runtime: Discovery Engine **`streamAssist` is grounded chat, not a
tool-calling/code-execution agent loop**. So we adopt the *ideas*, not the implementation.

## Decision

Adopt a **hybrid** context model. Keep our estate grounding; add a **working-document tool
surface** that treats the active document as a lazily-read, auditably-written environment.

**Layer A — Grounding (unchanged).** Discovery Engine search/grounding over the `UnitDescriptor`
(notebook + federated sources). This is our moat; leave it.

**Layer B — Working-document tool surface (new).** Five adopted elements:

1. **`<doc_state>` snapshot.** A compact, structured, **untrusted-wrapped** description of the
   active document injected each turn: surface, selection, inventory (sheets/slides/sections),
   named ranges, and comments. Schema in `@ge/contracts`; builder in `@ge/content`; produced by
   each bridge's `capture`.
2. **Lazy host-read tools + a client-side tool loop.** Expose a small, typed tool surface —
   `read_range(a1)`, `search_document(query)`, `get_outline()` — that the **`runtime`
   Orchestrator/AssistSession** drives in a bounded loop, fetching only what the turn needs
   instead of pre-chunking the whole document. The loop lives in our orchestrator because
   `streamAssist` will not run it for us.
3. **Formula-first writes.** When a result can be expressed as a host-native, user-inspectable
   formula, `bridge-excel` prefers emitting the **formula** over a literal value.
4. **Comments-as-citations.** Every externally-sourced write also drops a **source comment**
   (Excel cell comment / Word comment) carrying the citation, layered on top of — not replacing —
   our durable provenance metadata.
5. **Conversation compaction** in `AssistSession`: summarise older turns when the budget is
   approached, so long working sessions survive without losing the thread.

**Explicitly out of scope (skip).** We do **not** ship a Python/coreutils virtual filesystem in
the client. If true code-execution depth is needed (e.g. pandas over a large workbook), route it
to an **Agent Engine specialist agent via A2A** (per the standing preference over
`streamAssist.agentsSpec`) and give *that* agent a code-execution tool over the document
representation — keeping the heavy runtime server-side and the client thin.

## Consequences

- **`@ge/content` gains a `<doc_state>` builder** and a stable text/markdown rendering of blocks;
  the budgeter stays but is no longer the only path — large working docs are read on demand.
- **`@ge/runtime` gains a bounded tool-read loop and compaction** in the assist turn. The
  `DocBridge` interface grows narrow, typed read tools (`readRange`/`searchDocument`/`getOutline`)
  alongside the existing `listContext`/`resolveContext`/`actuate`/`watch`.
- **Bridges gain auditable-write discipline:** `bridge-excel` formula-first + cell source-comments;
  `bridge-word` keeps tracked changes and adds comment-anchored citations. Both still flow through
  the reversible **actuation gate** with full provenance (agent id, sources, identity, timestamp,
  content hash) — comments are an *additional*, human-visible citation, not the system of record.
- **Untrusted boundary is explicit:** `<doc_state>` and tool-read results are wrapped and screened
  (Model Armor at the engine) and passed to the model as data, never as instructions.
- **No change to identity or transport:** still client-direct, still the user's WIF-federated
  token (ADR-0001). Context construction and auth are independent layers.

## What stays
The `@ge/contracts` boundary; estate grounding on the `unit`; content-anchored Word findings
(`body.search` + apply-time re-resolve, now also used by `search_document`); provenance in host
metadata; reversible/tracked writes; identity-scoped reads; residency-pinned `discoveryengine`
endpoint. Our differentiators — estate grounding, durable cross-session provenance, end-to-end
identity federation, and five surfaces + Teams — are unchanged; Layer B closes the working-document
depth gap on top of them.
