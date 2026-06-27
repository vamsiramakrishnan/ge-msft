# Design — Gemini Enterprise for Microsoft 365

> **⚠️ Updated by `ADR-0001`.** This doc's **five experience invariants** (§2: the agent recedes,
> the unit travels, one identity envelope, provenanced + reversible, grounded-or-it-says-so) and the
> **per-surface verbs** (§4) remain the design north star and are still accurate. What changed: there
> is **no Surface Gateway** — wherever this doc says "one Surface Gateway," read "the shared
> surface-agnostic core (`runtime` + `web-shell`) calling Discovery Engine client-direct." The
> **phasing** (§6) is historical; **Outlook** is now a first-class sixth surface (not in the original
> five), and all six bridges are built. See `docs/STATUS.md` for current state.

**A multi-surface add-in: one agent, one research unit, one identity — across Word, Excel, PowerPoint, OneNote, Outlook, and Teams.**
*Phased design doc. Companion to the architecture and implementation docs, and the surface prototypes.*

---

## 1. North star

A knowledge worker shouldn't "use a Gemini app." They should find Gemini Enterprise *already present* in the Microsoft surface they're working in — reviewing the contract in Word, modelling the spend in Excel, building the QBR in PowerPoint, capturing research in OneNote, deciding in a Teams meeting — grounded the whole time on the *same* curated set of trusted sources, acting with *their* identity, leaving changes that are *traceable and reversible*.

The strategic point, carried from the surfaces plan: this is not six products. It is one shared
surface-agnostic core (`runtime` + `web-shell`) and one research unit, with six thin bridges. *(The
original framing said "one Surface Gateway" — there is no gateway now; see `ADR-0001`. The leverage
argument is unchanged: write the core once, add a thin bridge per surface.)* The clients differ only
where the *surface* differs; everything that constitutes the intelligence, the grounding, the
identity, and the trust is shared and built once.

```
                         ┌──────────────  THE RESEARCH UNIT  ──────────────┐
                         │  curated NotebookLM notebook  (precision/trust) │
                         │  + federated SharePoint/OneDrive (breadth/live) │
                         │  + the working surface          (what you touch)│
                         └───────────────────────┬─────────────────────────┘
                                                 │  one identity envelope
        ┌──────────┬──────────────┬──────────────┼──────────────┬──────────────┐
     ▼  Word     ▼ Excel        ▼ PowerPoint    ▼ OneNote       ▼ Teams
     editor /    analyst        composer        researcher      collaborator
     reviewer                                   / curator       / facilitator
        └──────────┴──────────────┴──────────────┴──────────────┴──────────────┘
                                  one Surface Gateway · one set of agents (A2A)
```

---

## 2. Experience principles (the invariants on every surface)

These five hold identically in Word, Excel, PowerPoint, OneNote, and Teams. If a surface breaks one, it's wrong.

1. **The agent recedes into the surface.** No bolted-on chatbot. The intelligence appears as the surface's *own* materials — inline annotations and tracked changes in Word, streaming cells and entity cards in Excel, slides and speaker notes in PowerPoint, page blocks and source tags in OneNote, live notes and action items in a Teams meeting. The side panel is for conversation and control; the *work* lands in the document.
2. **The unit travels.** The same curated notebook and the same federated connector sources ground the agent on every surface. Assemble the unit once; it follows you from the OneNote where you researched it to the Teams meeting where you decide on it.
3. **One identity envelope.** The user's Microsoft identity propagates to Gemini and back out to SharePoint/OneDrive via federated connectors and actions. The agent reads, grounds, and writes only what the user themselves can — on every surface, read and write.
4. **Every change is provenanced and reversible.** Edits arrive as tracked changes or citation-tagged blocks; each carries its agent, sources, and a content hash in the file's custom metadata. Nothing the agent does is silent or unattributable.
5. **Grounded or it says so.** In regulated contexts the agent answers from the unit and connected data or explicitly declines — it does not assert ungrounded. A claim without a source is a bug.

---

## 3. The unit (what grounds every surface)

Recap from the research-unit design, because it's the shared substrate: the unit is a *composed grounding scope* — a **NotebookLM notebook** as the curated, bounded, trusted core (precision; "answer only from these sources"), **federated SharePoint/OneDrive** connectors as the live, identity-scoped, no-copy edge (breadth and freshness on demand), and the **working surface** itself (the doc, workbook, deck, page, or transcript). Connector **actions** (upload/download/check-in-out) let the agent write results back to the Microsoft estate. The notebook is the backbone; the connectors are the elastic edges; the surface is what you're acting on.

Every surface in this design shows the unit in its panel and lets the user compose it — "add a notebook," "add docs." That panel control is the *same component* everywhere.

---

## 4. The experience, surface by surface

Each surface gets a verb — what the agent *is* there — and a signature interaction set that fits the surface's native materials.

| Surface | The agent is a… | Signature interactions |
|---|---|---|
| **Word** | editor / reviewer | Inline annotations (style, policy, verified) with grounded hover cards; comment thread as an agent task queue; surgical clause regeneration via content controls; provenance baked into the .docx |
| **Excel** | analyst | `=GE.ASK(prompt, range)` streaming custom function answering in-cell, grounded on the unit; vendor/entity **linked-entity cells** that expand into agent-enriched cards; "explain this range" |
| **PowerPoint** | composer | "Draft slides from the unit" streaming new slides in, every claim source-backed; speaker-notes generation; layout redesign suggestions |
| **OneNote** | researcher / curator | Source review of the notebook unit; "synthesise onto the page" with a citation tag per claim; **audio/video overviews** of the source set; discover-and-curate new sources |
| **Teams** | collaborator / facilitator | In-meeting live notes and grounded action items; ask-the-agent in meeting chat; a message extension to ground any message on the unit; a recap card posted to the channel or saved to OneNote |

The deliberate division of labour: **OneNote and Teams are where the unit is *assembled and decided on*; Word, Excel, and PowerPoint are where it is *acted on and produced from*.** Research and collaboration on one side, authored output on the other — the same unit flowing between them.

---

## 5. Cross-surface continuity (the thing competitors can't easily copy)

The magic isn't any single surface; it's the *flow* across them with state intact:

> Assemble the trusted sources in a **OneNote** notebook → the agent synthesises and makes an audio overview → carry the same unit into **Word** to draft and redline the contract → drop into **Excel** to model the liability exposure with `=GE.ASK` grounded on the same sources → build the **PowerPoint** QBR where every slide traces to the unit → walk into the **Teams** review where the agent already knows the unit and produces grounded action items → and the redlined doc is written back to SharePoint via connector actions.

One unit, one identity, one provenance trail, the whole way. Each handoff is free because the grounding scope and the identity envelope are shared infrastructure, not re-established per app. That continuity is the product.

---

## 6. Phasing

Five phases, each gated. The ordering retires the hardest risk first and front-loads the deepest surface, then reuses the spine outward.

**Phase 0 — Foundation.** Build the Surface Gateway, the unit resolver (notebook reference + federated connector resolution + actions), identity federation (Entra → Workforce Identity Federation), the provenance/signing service, and the shared web-app shell (the panel + unit composer reused by every surface). *Gate:* the slice-1 spine works end to end — a signed-in user, federated to Google, gets a grounded streamed answer over a selection.

**Phase 1 — Word.** The deepest surface and the one that proves the hardest UX. Ship inline annotations with grounded cards, the comment-as-task-queue, surgical regeneration, and provenance. *Exit:* a real review workflow run end to end with a design-partner customer (an FSI contract review is the obvious bellwether).

**Phase 2 — Excel + PowerPoint.** The "produce from the unit" surfaces. Net-new work is client surface code only — the gateway, unit, identity, and provenance are already built. Excel: streaming custom functions and linked-entity cells. PowerPoint: the deck composer and speaker notes. *Exit:* one analyst workflow (Excel) and one deck workflow (PowerPoint) grounded on the same unit a Word user assembled.

**Phase 3 — OneNote + Teams.** The "assemble and decide" surfaces, and the ones with the most divergent client models. OneNote: research capture and NotebookLM overviews — noting it ships as a *separate* package (web-only, legacy manifest). Teams: the meeting app, bot, and message extension, which need their own (bot) infrastructure. *Exit:* the cross-surface flow from §5 demonstrated unbroken across at least three surfaces.

**Phase 4 — Continuity & scale.** Cross-surface unit/provenance continuity hardened; admin governance and residency pinning per tenant; observability and eval across surfaces; distribution via AppSource / Teams Store / centralized admin deployment. Then bellwether-and-replicate to additional tenants and verticals. *Exit:* GA across surfaces with an admin able to turn it on safely for a large seat count.

```
 P0 Foundation ──► P1 Word ──► P2 Excel+PPT ──► P3 OneNote+Teams ──► P4 Continuity+Scale
   gateway,         deepest      reuse spine,      divergent           one unit travels;
   unit, identity,  surface;     surface code      client models;      governance, AppSource,
   provenance       bellwether   only              OneNote/bot infra   replicate
```

---

## 7. How we'll know it's working

- **Continuity rate** — share of sessions where a unit assembled on one surface is reused on another without re-establishing it. *This is the number that proves the thesis.*
- **Reuse ratio** — share of each new surface served by the shared gateway/unit/identity vs bespoke code. Should rise toward 1 as surfaces are added.
- **Time-to-surface** — adding a new surface should take surface-client time (days), not a new backend (weeks).
- **Grounded-answer rate** — share of agent assertions carrying a citation to the unit or connected data; ungrounded assertions in regulated contexts trend to zero.
- **Identity integrity** — share of read/write operations correctly scoped to the signed-in user.
- **Adoption per surface** — MAU and task completion, by surface and by the agent's verb (review, analyse, compose, research, decide).

---

## 8. Risks specific to going multi-surface

- **OneNote is the odd one out** — web-only, legacy XML manifest, thinner API. It ships as a companion package, not part of the unified bundle, and its capabilities are narrower. Scope its experience to what the API actually supports.
- **Teams needs more than a panel** — the bot and message-extension capabilities require Bot Framework / Azure Bot infrastructure and their own consent model. It's the heaviest client.
- **Marketplace review is a gated, multi-week path** — AppSource and Teams Store certification (and the cross-platform "works everywhere you declare it" requirement) start early, not at launch.
- **One unit, many ACLs** — the unit must enforce the *intersection* of what the user can see across the notebook and every connector; the identity envelope is the control, and it's the part a security review will probe hardest.
- **Surface drift** — resist letting each surface grow its own divergent agent behaviour. The verbs differ; the agent, the unit, and the trust model do not.

---

## 9. One sentence

Define the agent once by what it grounds on (the unit) and how it behaves (the five invariants), then express it through each surface's native materials — annotations in Word, cells in Excel, slides in PowerPoint, page synthesis in OneNote, meeting notes in Teams — so that one identity-scoped, fully-provenanced research unit flows with the user across the entire Microsoft suite.
