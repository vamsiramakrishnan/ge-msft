# EXPERIENCE.md — The Gemini command surface for Microsoft 365

> **One-sentence promise:** *Ask Gemini anything about a scope you choose, grounded on sources you assembled, acting as you — and see and undo everything it touches.*

This is the reimagined human-facing layer over the capability stack (ADR-0003→0006). The capability stack underneath — the `cmd` algebra, the plan→approve→gate loop, the bridges, the closure checker — **does not change**. What changes is the *vocabulary* and *surfacing*: the verbs stop being contract-review task names and become Copilot-altitude **general capabilities**, scope becomes a first-class orthogonal axis, and a custom prompt over any scope becomes a first-class primitive. The differentiators (client-direct identity, reversible+provenanced writes, the legible plan, the GE skills) become *how each general verb behaves*, not a separate vertical vocabulary.

---

## 1. The capability model — verb × scope × ground

The user flagged the core defect correctly: `regen-clause`, `resolve-comment`, `draft-slides`, `synthesize`, `meeting-notes` are **tasks from the contract-review demo smuggled into the capability type system**. The proof is in the repo: `command-palette.ts:48` already renames `regen-clause` → `/rewrite`, and `intent.ts:11` literally documents it as "rewrite one content control" — leaking Word OOXML into a cross-surface boundary type. Three genuinely generative quick actions (`write-formula`, `risk-column`, `draft-reply`) are crammed into `assist` because there is no general generate/rewrite verb.

### The decision: a two-tier model — small GENERAL verb set + orthogonal SCOPE + orthogonal GROUND

**Tier 1 — the verb (WHAT, general, Copilot-shaped). The final `IntentSchema` is exactly seven:**

| Verb | `/label` | Meaning | Output | Route |
|---|---|---|---|---|
| `ask` | `/ask` | Grounded chat / **a custom free-text prompt** over a scope (the rename of `assist`) | chat | `send` |
| `summarize` | `/summarize` | Condense the scope | chat | `send` |
| `explain` | `/explain` | Clarify the scope in plain language | chat | `send` |
| `rewrite` | `/rewrite` | Apply **any instruction** to the scope → a **reversible** edit (absorbs `regen-clause` + all in-place text/cell/slide edits) | write | `runCommands` |
| `review` | `/review` | Whole-scope pass emitting N findings → N gated annotations | annotation | `runCommands` |
| `draft` | `/draft` | Generate **new** material from the unit (slides, a OneNote page, a reply, a column) — absorbs `draft-slides`, `synthesize`, `save-to-onenote`, `compose` | write | `runCommands` |
| `notes` | `/notes` | Transcript → live notes + action items (Teams; the one verb whose closure is a meeting transcript) | annotation | `runCommands` |

**Why these four "specialist" verbs (`rewrite`, `review`, `draft`, `notes`) stay distinct and the others collapse:** the structural test is **"does it fan out to a multi-step plan a generic single-shot can't express?"** `ask`/`summarize`/`explain` are single-shot reads → chat. `rewrite` lands a reversible write. `review` fans out to N findings (the plan IS the audit artifact). `draft` generates new host material. `notes` is transcript→annotations. `resolve-comment` is **deleted as a verb** — it becomes `rewrite` or `review` with `scope: comment(id)`, the same code path. `synthesize`/`meeting-notes`/`draft-slides` are **deleted** — they were surface-bound closures of `draft`/`notes`.

**`rewrite` is the load-bearing generalization.** It carries a free-text `instruction` + a `scope`, and compiles to whatever reversible write the scope×surface affords: `suggest "old"=>"new"` (Word tracked change), `set`/`format` (Excel cell), slide-body replace (PPT). "Tighten", "make formal", "rewrite this clause to match the policy" all become `rewrite` + an instruction string — exactly Copilot's "Rewrite with a prompt," except it lands as a tracked change through the gate.

**Tier 2 — SCOPE (WHERE, orthogonal). A new first-class field, never a verb.**

```
CommandScope = selection | document | range(<A1|named>) | section(<heading>)
             | comment(<id>) | this-item
```

Every verb takes a scope. `command-plan.ts` already has a `scope?: string` field threaded through the parser (`:28`, `:157`, `:196`) but **the capture surface can't reach it** — it's smuggled via the magic `@this`/`@unit` mention strings. We surface scope as a **segmented control next to Send** ("Selection ▾ / Document"), defaulting per surface, with surface-named labels supplied as **data** in the palette (so `web-shell` stays surface-agnostic — see §4). Right-click hard-binds `selection`; quick actions default it.

**Tier 3 — GROUND (WHAT IT'S GROUNDED ON, orthogonal).** The `@`-mention picker output: the research unit, a notebook, a `datastore`, `this`, an upload. `ground` and `mentionKinds` are **the same concept** (the audit confirms: `quick-action-seed.ts:11` renders each ground token as `@token`) and become one typed `GroundSourceSchema`.

The result: **`rewrite scope:section(§4) ground:@VendorRiskPolicy`** is a legible one-liner — *rewrite this section to comply with that source* — a sentence Copilot cannot express because its grounding is the implicit Graph and its scope is the implicit selection.

### Mapping onto the existing layers

- **`IntentSchema`** → the 7-verb enum above. `INTENT_REQUIRES` re-keyed: `rewrite` requires `tracked-change|replace-selection|write-cells|fill-content-control` (closure now reads "can this surface land a reversible edit at this scope"); `draft` requires `insert-slide|append-page|create-mail|reply-mail`; `ask`/`summarize`/`explain` require nothing.
- **`CommandPlan`** → already has `scope`; add `scope: CommandScopeSchema` typing (was a bare string) and re-key `intent`.
- **`QuickAction`** → gains a typed `scope` field; its `output` becomes **derivable** from intent (write/annotation ⇒ non-empty `INTENT_REQUIRES`), closing the silent drift the audit found in `draft-reply`/`risk-column`/`write-formula`.
- **The `cmd` executor grammar** (`set/suggest/comment/slide/page/mail/post/compose/format/reply`) is **UNCHANGED** — it was already general. Only the human-facing intent tier moves. The executor maps general-verb × surface → the existing `WRITE_VERB_TO_KIND` kinds.
- **The GE skills** (`parse_plan.py`, `parse_commands.py`) mirror the TS — they move in the same wave (the TS side is authoritative; a parity test pins them).

---

## 2. Per-surface canonical flows

Each surface should feel made of its own materials. The verb set is identical everywhere; the **presets** (saved verb×scope×ground×instruction tuples) stay rich per surface, and the **landing** is surface-native and reversible.

**Word** — entry: pane + ribbon "Ask Gemini" + right-click "Ask Gemini about this".
- `/review scope:document ground:@<picked source>` → findings land as content-anchored tracked changes, each with a grounded hover card (agent id, source, hash) and Accept/Dismiss. Anchored by `body.search`, re-resolved at apply-time, **degrades to a panel item on drift** (never a broken anchor).
- select text → `/rewrite "tighten, preserve meaning"` → tracked change, reversible. Where a content control exists, `rewrite scope:section` compiles to `fill-content-control`; otherwise a coarse tracked change (the per-surface scope→actuation mapping is explicit — see risks).
- comment thread as task queue → `/rewrite scope:comment(id)` edits the anchored text + replies + resolves.
- **Catalog fix:** "Review against policy"/"Find unsupported claims" recut to a general **"Review against…"** preset that takes an `@source` (policy is just one possible `@doc`). The contract-review nouns move to an **optional vertical pack**, not the default catalog.

**Excel** — entry: pane + ribbon + `=GE.ASK(prompt, range)` in the grid + right-click on a range.
- `=GE.ASK(...)` streams a grounded answer into the cell, fill-down across rows.
- select range → `/summarize` (totals/trends/outliers) → chat with one-click "write as column" landing address-anchored, provenanced cells past `isUnsafeFormula` + gate.
- `/rewrite scope:range "add a derived column that flags …"` (generalizes "risk-column"); **`=GE.ASK` becomes a first-class quick action** (it isn't in the catalog today). Excel gets `/rewrite` and `/review` wired (today the palette claims `[assist, review]` but ships generative actions mislabeled `assist`).

**PowerPoint** — entry: pane + ribbon + right-click on a slide.
- `/draft scope:deck ground:@unit "the risk section"` → source-backed slides stream in, each bullet traceable, inserted as new slides (reversible by deletion + provenance).
- select slide → `/summarize scope:deck` or `/ask "redesign this slide"` → chat.
- **Catalog fix:** "Generate speaker notes" is a **phantom** (CAPABILITY-MAP: `set-speaker-notes` is modeled-not-advertised, no host write path). Demote to `output:'chat'` (draft notes *into the pane*) until the bridge can actuate.

**OneNote** — entry: **pane only** (web-only, legacy manifest — no ribbon/right-click parity; the design must not assume uniform entry points).
- `/draft scope:page ground:@unit` → a cited summary block appended to the page, one **inline citation tag per claim** (NotebookLM-grade "answer only from these sources" made visible).
- **OneNote is where the unit is assembled** — add a **"add this page's sources to the unit"** composition action (its signature job, currently unserved by the catalog).
- "audio overview"/"discover sources" stay `output:'chat'`.

**Outlook** — entry: pane + right-click on a message (the staged reply is the write; no inline).
- `/summarize scope:thread` ("catch me up").
- `/draft "a reply" scope:thread` → a **staged reply** (`displayReplyForm`), never auto-sent, editable before send, with **tone control** (Copilot has tone; we should expose it). The on-send gate is the reversibility story for mail.
- **Scope toggle** (this message vs whole thread) is exposed explicitly (today it's conflated in prompt text). Add `mail` to the manifest `contextMenus` scopes so Outlook gets a right-click menu at all.

**Teams** — entry: meeting side-panel (Live notes / Ask tabs) + message extension.
- `/notes` streams live notes during the meeting, grounded on `@unit`+transcript, action items with owners.
- `/ask scope:last-5-min` or `scope:transcript` → in-meeting Q&A.
- recap → "Post to channel"/"Save to OneNote", both **reviewable before posting**. **Honesty:** Teams has no durable provenance write yet (STATUS) — label these "staged/reviewable," not "provenanced," until the host-metadata write lands. Transcript scope (rolling window vs whole) is a real toggle.

---

## 3. How `/` + `@` + buttons + right-click + plan-gate cohere

**One composer, one dispatch, one lifecycle.** Today the pane is four parallel entry points wired through two opaque routes with the gate decision duplicated (`onQuickAction` routes on `output`; `onInvoke` routes on `intent !== 'assist'` — they can drift). The fix:

1. **Buttons and right-click PRE-FILL the composer** rather than dispatching a hidden path. A chip types its `verb @ground scope instruction` into the box and submits — so the user *sees* "the same thing you could have typed" and learns the grammar by watching. Right-click on a comment/range/slide pre-fills the composer scoped to that target.
2. **One `Invocation` type → one controller dispatch.** `{ intent, scope, mentions: {kind, ref}[], instruction }` — typed, not a re-parsed raw string. App stops making the gate decision twice; one `isActuating(intent)` predicate (derived from `output`/`INTENT_REQUIRES`) routes: chat verbs → `send`, write/annotation verbs → `runCommands`. The route is **total and tested** (the audit's key risk: a `rewrite` must *never* reach `send`).
3. **Delete the "Agentic (read & propose writes)" checkbox.** Mode is **inferred**: any verb that writes routes through the plan→approve→gate loop; a bare question routes to grounded chat. The user never picks "chat vs agent." Fail-closed: if the planner emits *any* effect, gate; an ambiguous turn asks a `clarify` rather than guessing.
4. **One five-step turn lifecycle for every route:** (1) Grounding — the unit chips that fed this turn light up; (2) Streaming answer + inline citation pills; (3) Plan — IF it writes, the dry-run effect-set as a confirmable program (verbatim `cmd` lines, reuse `PlanApprovalCard`); (4) Confirm — one gate; (5) Landed — effects drop into a persistent **Changes ledger** with per-change Undo (verb, target, sources, identity, timestamp, content-hash). Read-only turns stop at step 2. A free-form `/rewrite` lands through the same 5 steps as a prebuilt button.
5. **The research unit is a persistent top bar** above the thread on every surface — notebook + federated sources + "this {selection|range|slide|thread}" as chips, with +Add. `@`-mentions ADD to this bar. Scope is always on screen, so grounding is *felt before* the answer streams. (Ledger and unit bar are collapsible in narrow panes — Outlook/Teams.)
6. **First-run = the unit, not a chat prompt.** Empty pane shows the Unit bar + "Add your first source" + 2-3 surface starter chips, under the promise line: *"Grounded on sources you choose. Acting as you. Every change reviewable."*

---

## 4. Keeping the core surface-agnostic

Surface-specific scope labels ("range | sheet | table" vs "slide | deck") live as **data in `CommandPaletteSpec`** (a `scopeOptions` list per surface), consumed by the surface-agnostic `Composer`. No host-specifics leak into `web-shell`. The general/specialized catalog split is structural: `UNIVERSAL_ACTIONS` (generated for every `Surface` from `SurfaceSchema.options`) + `SURFACE_ACTIONS` (genuine specializations); adding a surface needs no edit to the universal block.

---

## 5. What this unlocks beyond Copilot

- **Reversible Rewrite.** Copilot's Rewrite mutates text in place; ours stages every `rewrite` as a tracked change carrying the instruction, sources, identity, and content hash — diff visible, accept/reject per-change, the `.docx` remembers who/what/why. "Rewrite with a prompt" becomes auditable.
- **Scope-as-grounding fusion.** `rewrite scope:section(§4) ground:@VendorRiskPolicy` is a legible sentence Copilot can't say.
- **Preset transparency.** A quick action / right-click item is literally a tuple the user can SEE and edit before it runs — versus Copilot's opaque one-click prompts. The same tuple is what the planner emits and the executor runs: one legible artifact end to end.
- **Custom prompt over any scope, gated.** `ask`/`rewrite` take an arbitrary instruction against selection OR document OR a named range/section/comment — matching Copilot's "do X to this," but the write half is gated and provenanced, so a regulated tenant can let users run free-text edits safely.
- **Specialist fan-out stays legible.** `review scope:document` shows as a PLAN (N proposed findings) you approve once, then lands N individually-reversible annotations. Copilot's agentic review applies silently.
- **The unit travels.** Assemble once in OneNote; the same bounded source set grounds the Word redline, the Excel model, the PPT QBR, the Teams decision. Copilot's Work IQ is one ambient Graph the user can't compose or bound.
