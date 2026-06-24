# Microsoft 365 Copilot — what it does, and how we differ

**The incumbent in every surface we ship into.** This doc is the competitive baseline: what
Microsoft 365 Copilot actually does per app, how it grounds, how it extends — and where our
client-direct Gemini Enterprise add-in is genuinely different (not just differently branded). It is a
companion to `02-design.md` (our five experience invariants) and `CAPABILITY-MAP.md` (our honest I/O
inventory). Read it before pitching, demoing, or scoping a surface against Copilot.

The short version: Copilot is **native, opaque, and edits directly**. We are **a sideloaded add-in,
legible, and write reversibly with provenance** — running *Gemini* inside the Microsoft estate with
the user's own identity federated to Google. Same surfaces, different trust posture.

---

## 1. Delivery model — native platform vs sideloaded add-in

This is the first and largest asymmetry, and it cuts both ways.

**Copilot is native.** It ships as part of Microsoft 365 itself, in three shapes:

- **In-app pane** — the Copilot side pane *inside* Word, Excel, PowerPoint, Outlook, OneNote, and
  Loop, plus inline entry points (the Copilot icon in the margin/ribbon, draft-with-Copilot prompts).
- **Standalone Copilot Chat** — the BizChat / `microsoft365.com/chat` web and Teams app, grounded on
  the whole Microsoft Graph, separate from any one document.
- **An agent platform** — Copilot Studio + declarative/custom agents that appear across the suite and
  in Teams, the agent store, and the Copilot pane itself.

As of 2026, Copilot's **agentic (plan / execute / refine) mode is GA in Word, Excel, and
PowerPoint** — Copilot will plan a multi-step task, run it across the document, and iterate, not just
answer a single prompt. It is the first-party intelligence layer of the suite, with no install step
and privileged in-app entry points we cannot match.

**We are a sideloaded add-in.** We bring *Gemini Enterprise* into the same surfaces through the
public extensibility contract: a unified M365 manifest (Package A) plus a legacy OneNote manifest
(Package B), an Office.js / TeamsJS task pane, distributed via AppSource or M365 admin push. We do
not get native inline entry points; we get the add-in task pane and the command surface inside it.
What we get *instead* is the thing Microsoft structurally can't offer in its own first-party product:
**Gemini, the user's curated NotebookLM research unit, and a client-direct identity bridge from Entra
to Google** — see §6.

| | **Copilot** | **Us (GE for M365)** |
|---|---|---|
| Install | Native, part of M365 | Sideloaded add-in (AppSource / admin) |
| Entry points | In-app pane + inline + ribbon + standalone Chat | Add-in task pane (one per surface) |
| Intelligence | Microsoft's models | Gemini Enterprise (Discovery Engine `v1alpha`) |
| Agentic mode | GA in Word/Excel/PPT (plan/execute/refine) | Our `/ + @ + plan` command surface (ADR-0004/0005) |
| Identity to model | Microsoft-internal | Entra → Google WIF, client-direct, in the browser |
| Surfaces | Word, Excel, PPT, Outlook, OneNote, Teams, Loop | Word, Excel, PPT, Outlook, OneNote, Teams |

---

## 2. What Copilot does, app by app

A functional baseline. Each cell is "what a user expects Copilot to do here today," so we can hold
our per-surface verbs against it.

| App | What Copilot does |
|---|---|
| **Word** | Draft a document from a prompt or reference files; rewrite / summarise / shorten selections; ask questions about the doc; **agentic** multi-step authoring (plan a section set, draft, refine across the doc); inline coaching and tracked-suggestion-style edits |
| **Excel** | Generate formulas and explain them; surface insights, trends, and PivotTable/chart suggestions over a table; "highlight / sort / filter" by natural language; **agentic** analysis passes; Python-in-Excel assistance |
| **PowerPoint** | Build a deck from a prompt or from a Word doc; add/redesign slides; generate and refine speaker notes; summarise a deck; **agentic** deck construction with brand/template awareness |
| **Outlook** | Summarise long mail threads ("catch-up"); draft and refine replies with tone control; "coaching" on a draft; triage / prioritise the inbox; meeting-prep summaries from related mail |
| **OneNote** | Summarise pages and notebooks; draft notes, plans, and to-do lists; rewrite/reformat; ask across notebook content |
| **Teams** | Live meeting recap, notes, and action items; "catch me up" on chats and channels; intelligent recap of recorded meetings; draft messages; meeting Q&A grounded on the transcript |
| **Loop** | Draft and summarise Loop pages/components; ideate and co-create in the live collaborative canvas; carry workspace context into the Copilot pane |

Note Copilot's **Loop** surface has no analog in our build — we ship six surfaces and Loop is not one
of them. Conversely, our cross-surface continuity (the *same* research unit travelling from OneNote to
Teams to Word, §5 of `02-design.md`) is not something Copilot foregrounds, because its grounding is
the implicit Graph, not a unit the user composes (§3 below).

---

## 3. Grounding — Work IQ / Graph vs our two planes

Copilot grounds on the **Microsoft Graph**, marketed under "Work IQ": the user's mail, files, chats,
meetings, calendar, and people, plus the **currently open document** and, when enabled, the **web**.
It is automatic and ambient — the user does not assemble a source set; Copilot reaches into whatever
the Graph says they can see, scoped by their Microsoft permissions.

Our grounding is **explicit and two-planed**, and it maps cleanly onto Copilot's surface:

| Copilot grounds on… | …maps to our |
|---|---|
| The open document | **Plane A** — in-document capture via the bridge (Office.js / TeamsJS) |
| A composed source set (Copilot has none; it's implicit) | **Plane A** — the **research unit**: curated NotebookLM notebook + federated SharePoint/OneDrive + the working surface |
| Microsoft Graph / Work IQ (mail, files, chats, meetings, people) | **Plane B** — the **estate**, read delegated via `@ge/graph-client` (`/search/query`, `/me/messages`, `/me/events`, `/users`) and grounded via GE federated connectors |
| Web | GE engine-side web grounding (engine config) |

Two real differences fall out of this:

1. **The unit is composable; Work IQ is not.** A user explicitly assembles the research unit
   ("add a notebook," "add docs") and the *same* unit grounds every surface — precision (NotebookLM:
   "answer only from these sources") plus live breadth (federated connectors), versus Copilot's
   one ambient Graph. This is `02-design.md`'s "the unit travels" invariant; Copilot has no equivalent
   curated-and-bounded mode.
2. **Plane A vs Plane B is an explicit seam for us.** Copilot blurs open-doc and estate into one
   reach. We keep them separate — in-document content is captured natively and framed strictly as
   *data* (never as instructions; `CAPABILITY-MAP.md` "untrusted input"), and the estate is a distinct,
   delegated, identity-scoped read.

---

## 4. Extensibility — the analog table

Both stacks let a builder extend the assistant. The shapes line up almost one-to-one, which is useful:
it means the patterns a Copilot Studio author already knows translate to our skill surface.

| Copilot extensibility | Our analog | Notes |
|---|---|---|
| **Declarative agents** (Copilot Studio) — instructions + grounding + starter prompts | **GE skill** — a GE *agent* with a `skillAgentDefinition` (`instruction` + `subfiles[]`), mounted per turn via **`skillsSpec`** | Our `m365-surface-commander` (executor) and `m365-command-planner` (plan builder) are exactly this. See `docs/api/discoveryengine/skills-and-agents.md`. |
| **Graph connectors** — index external content into the Graph for Copilot to ground on | **Federated connectors / `dataStoreSpecs`** — the `@` mention surface; the engine queries SharePoint/OneDrive *as the user*, no-copy | We prefer **federated** mode for ad-hoc sources, ingestion only for large stable corpora (CLAUDE.md constraint). |
| **Plugins / actions** (API-backed skills, message extensions) | **The `cmd` algebra** (ADR-0004/0005) + **A2A specialist agents** on Agent Engine | Actuation is our composable command grammar; specialist work routes over A2A, *not* StreamAssist's `agentsSpec` (known agent-id bug — CLAUDE.md). |
| **Custom engine agents** (full Copilot Studio orchestration) | **A2A agents** (Review, Redline, …) reached directly on Agent Engine | StreamAssist is reserved for the grounded-assistant chat path; specialists route around it. |

The headline: a **Copilot Studio declarative agent ≈ a GE skill bundle**, and a **Graph connector ≈ a
GE federated connector**. The substantive divergence is at the action layer — Copilot plugins call out
to APIs opaquely; our actions are the *legible* `cmd` algebra the planner emits and the executor runs,
and they land as reversible writes (§6).

---

## 5. Parity map — Copilot verbs to our buttons / IntentSchema

This ties Copilot's per-app expectations to our prebuilt buttons and the seven general
`IntentSchema` verbs (`packages/contracts/src/intent.ts`: `ask`, `summarize`, `explain`, `rewrite`,
`review`, `draft`, `notes` — see `docs/EXPERIENCE.md`). The old surface-bound task names
(`resolve-comment`, `regen-clause`, `draft-slides`, `synthesize`, `meeting-notes`) collapse into a
general verb plus an orthogonal **scope** (`selection | document | range | section | comment |
this-item`). It is the "yes, we do that too" table — and where we do it *differently*, the difference
is the pitch.

| Surface | Copilot verb | Our prebuilt button / interaction | Verb × scope |
|---|---|---|---|
| **Word** | Summarise / ask about the doc | Ask-about-selection; grounded hover cards | `ask` / `summarize` · `scope:selection` |
| **Word** | Rewrite / suggest edits | Inline review → findings as **tracked changes**; "Accept change" | `review` · `scope:document` |
| **Word** | Resolve a comment | Comment thread as agent task queue; edit + reply + resolve | `rewrite` · `scope:comment(id)` |
| **Word** | Rewrite this clause | Surgical clause regeneration (tracked change / content control) | `rewrite` · `scope:section` |
| **Excel** | Generate formula / explain range | `=GE.ASK(prompt, range)` streaming custom function; "explain this range" | `ask` / `explain` · `scope:range` |
| **Excel** | Insights / highlight / sort | Linked-entity cells → enriched cards; `format` actuation | `summarize` · `scope:range` |
| **Excel** | Add a derived/risk column | "Add a column that flags …" → address-anchored, gated cells | `rewrite` · `scope:range` |
| **PowerPoint** | Build a deck / add slides | "Draft the risk section" → streamed, source-backed slides | `draft` · `scope:deck` |
| **PowerPoint** | Speaker notes | "Generate speaker notes" (drafts into pane; un-advertised host write per ADR-0006) | `draft` · `output:chat` |
| **Outlook** | Catch-up on a thread | Whole-item capture → grounded summary | `summarize` · `scope:this-item` |
| **Outlook** | Draft / refine a reply | Reviewable reply (`displayReplyForm`), staged not sent, with tone control | `draft` · `scope:this-item` |
| **OneNote** | Summarise onto the page | "Summarise sources onto page" with a citation tag per claim | `draft` · `scope:section` |
| **OneNote** | Audio/video overview | "Make an audio overview" of the source set | `ask` · `output:chat` |
| **Teams** | Live notes / action items | In-meeting live notes + grounded action items | `notes` · `scope:this-item` |
| **Teams** | Recap to channel / OneNote | Recap card "Post to channel" / "Save to OneNote" (reviewable) | `draft` · `scope:this-item` |

Two parity caveats worth stating honestly: we **do not** ship a Loop surface, and our PowerPoint
speaker-notes and several Outlook/Graph write verbs are *modeled but not advertised* (ADR-0006 capability
closure: never advertise a verb the bridge can't actuate). Copilot advertises broadly and degrades
silently; we advertise only the closure of what we can actually do.

---

## 6. Where we differ — the pitch

Three differentiators, each a deliberate trade against Copilot's native advantage.

**1. Client-direct Entra → Google identity federation.** This is the whole reason the product can
exist: the add-in federates the *signed-in user's* Entra identity to Google via Workforce Identity
Federation **in the browser**, and calls Discovery Engine directly. Gemini runs inside M365 acting as
the user — reading, grounding, and writing only what that user can — **with no gateway holding
credentials and no Google secret ever reaching a client** (ADR-0001). Copilot cannot bring Gemini; we
can, and we do it under one identity envelope end to end. The client holds only a short-lived Entra
token and the federated Google access token derived from it, in memory.

**2. Reversible, provenanced writes vs direct agentic edits.** Copilot's GA agentic mode **edits the
document directly** — it plans and applies. Our writes are *reversible by construction*: findings land
as **tracked changes** (Word/PPT), **citation-tagged blocks** (OneNote/Teams), or **staged drafts**
(Outlook reply/compose, never auto-sent), and **every** change carries its agent id, sources, identity,
timestamp, and a **content hash** in the host's durable metadata. Word findings are anchored by content
(`body.search`) and **re-resolved at apply-time** — a stale finding degrades to a panel item rather than
a broken annotation. Nothing the agent does is silent or unattributable; "every change is provenanced
and reversible" is an invariant (`02-design.md` §2), not a feature flag. The actuation gate (`@ge/triggers`)
means a write requires confirmation; Copilot's plan/execute does not.

**3. A legible `/ + @ + plan` command surface vs "it just does it."** Copilot's agentic mode is
deliberately opaque — you prompt, it plans and acts behind the pane. Our command surface is **explicit**:
`/` verbs and `@` source mentions compile, via the planner skill, into a **confirmable `plan` block**
(intent / surface / scope / ground / step / exclude / clarify / confidence — see
`skill/m365-command-planner`), which the executor turns into the **`cmd` algebra** (ADR-0004) that
actuates. The user sees the plan before it runs, sees which sources ground it (`@`), sees which skill
ran (`invokedSkills[]` for provenance), and can edit any step. Where Copilot optimises for "magic," we
optimise for **legibility under audit** — which is the posture a regulated tenant actually needs.

The summary trade: Copilot wins on native reach and zero-install ubiquity. We win on **whose model,
whose identity, whose sources, and whether the user can see and reverse what the agent did.**

---

## Sources

Microsoft's own documentation and product blog, as the authoritative baseline (verify current state
before quoting specifics in a customer-facing deck — Copilot ships fast):

- Microsoft Learn — *Microsoft 365 Copilot overview* and per-app articles (Word, Excel, PowerPoint,
  Outlook, OneNote, Teams, Loop): <https://learn.microsoft.com/en-us/copilot/microsoft-365/>
- Microsoft Learn — *Copilot agentic capabilities / agent mode in Word, Excel, PowerPoint*:
  <https://learn.microsoft.com/en-us/microsoft-365-copilot/>
- Microsoft Learn — *Copilot Studio: declarative and custom engine agents*:
  <https://learn.microsoft.com/en-us/microsoft-copilot-studio/>
- Microsoft Learn — *Microsoft Graph connectors for Copilot*:
  <https://learn.microsoft.com/en-us/microsoftsearch/connectors-overview>
- Microsoft Learn — *Work IQ / Microsoft 365 Copilot and the Microsoft Graph*:
  <https://learn.microsoft.com/en-us/graph/copilot-concept-overview>
- Microsoft 365 blog — *What's new in Microsoft 365 Copilot* (agentic mode GA, surface rollouts):
  <https://www.microsoft.com/en-us/microsoft-365/blog/>

**Cross-links (our side):** `docs/api/discoveryengine/skills-and-agents.md` (the live skill lifecycle
— `skillAgentDefinition`, `files:upload`, the per-turn `skillsSpec` mount), `02-design.md` (experience
invariants + per-surface verbs), `CAPABILITY-MAP.md` (two-plane I/O inventory), `ADR-0001`
(client-direct), `ADR-0004`/`ADR-0005` (command protocol + composable capabilities), `ADR-0006`
(capability closure).
