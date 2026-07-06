# Task-pane UI

The view layer of the surface-agnostic web-shell: the React task pane that renders over
`PanelController` state. No Office.js / TeamsJS / Graph here — the bridges own host code, the
controller owns all state, network and approval logic. This doc covers how to see the panel, what
each component renders, and the design-token system. All paths are under `packages/web-shell`.

## Running the preview

```bash
npm run preview -w packages/web-shell
```

This starts a plain-HTTP Vite dev server on **http://localhost:3100** and opens `preview.html` —
no Office host, no TLS, no network. It mounts the **real** `<App/>` over a fake `PanelController`
(`makeMockController`) driven by scripted fixtures, so the whole panel renders in any browser tab.
Buttons log to the console instead of actuating.

The left toolbar drives the panel:

- **Surface** — switches `data-surface` (word / excel / powerpoint / outlook / teams), which swaps
  the per-surface host accent (focus ring + links).
- **Cards** — independent toggles for each state slice (Context, Thread, Suggestions, **Skills**,
  Run steps, Plan, Write, Proposals, Error, Busy), plus **All on** / **Idle / empty** presets. The
  **Skills** toggle shows/hides the `SkillsPanel` surface.

The preview is view-only and never ships in the add-in bundle (`vite.preview.config.ts` is separate
from `vite.config.ts`).

## Component map

The panel hierarchy, top to bottom (`src/taskpane/components/`):

| Component               | Renders                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          | Controller methods                                                                       |
| ----------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------- |
| `App`                   | Panel shell: `data-surface`, `aria-busy`, header (agent identity), and the `<main>` thread region landmark. Wires every child to the controller.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | `refreshContext`, `attach`, `detach` (via `onToggle`), `onAutomate`, `dismissSuggestion` |
| `ContextTray`           | The research-unit grounding scope as chips, split into "Attached sources" / "Available to attach" `role=list` groups; each chip is an attach/detach `<button>`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                  | `refreshContext`, `attach`, `detach`                                                     |
| suggestions `<section>` | Clickable suggestion cards (in `App`).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           | `onAutomate`, `dismissSuggestion`                                                        |
| `SkillsPanel`           | The in-session skills surface (ADR-0005 `def`), a `role=region` ("Skills") that lists each registered skill as a card: its `name(params)` signature, a "✓ registered" badge, the verbatim `def` line, bindable param `<input>`s (label-associated via `htmlFor`/`id`, prefilled from each param's example), and an "Invoke skill" `<button>`. Invoking does NOT actuate here — it routes through the agentic loop so the plan still lands on `PlanApprovalCard`. Renders nothing when no skills are registered.                                                                                                                  | `invokeSkill` (read-only view of `skills`)                                               |
| `MessageThread`         | The grounded conversation (`role=log`, `aria-live=polite`): user/assistant bubbles, streaming caret, a "Sources" eyebrow + citation pills. Each citation pill is now a `<button>` (`aria-expanded`/`aria-controls`) that opens a **source-detail popover** (title + locator + http(s)-only link) so the source is inspectable without leaving the thread.                                                                                                                                                                                                                                                                        | — (read-only view of `messages`)                                                         |
| `ProvenanceDetail`      | The provenance drill-down for an applied write, rendered as a `<dl>` (`aria-label="Change provenance"`): agent, signed-in identity, timestamp (`<time>`), grounding sources (http(s)-only links), and content hash. Presentational only.                                                                                                                                                                                                                                                                                                                                                                                         | — (read-only view of `Proposal.provenance`)                                              |
| `RunSteps`              | The command-loop transcript: an "Activity" section with an ordered, polite live region of run steps.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             | — (read-only view of `steps`)                                                            |
| `PlanApprovalCard`      | Fail-closed plan-level gate (ADR-0005): a `role=region`/`aria-live` card listing each effect's **verbatim** command line via `renderCommandLine`, with Approve / Reject. Each effect row is now an **expandable** `<button>` (`aria-expanded`/`aria-controls`) that reveals the dry-run review — the resolved target, the value it resolves to, and a before→after preview — so the whole plan is reviewable before it runs.                                                                                                                                                                                                     | `approvePlan`, `rejectPlan`                                                              |
| `WriteApprovalCard`     | Fail-closed per-write gate (ADR-0004): a `role=region`/`aria-live` card showing the **verbatim** `pending.command`, with Approve / Reject.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       | `approvePendingWrite`, `rejectPendingWrite`                                              |
| `ProposalCard`          | The reversible-write review: one card per staged proposal in its status; pending exposes an "Accept change" `<button>`. The body is now rendered **surface-faithfully** — Excel `write-cells` shows the value formula-first against its range target (with an optional `◆` linked-entity card: gradient header, key/value rows, "loaded from the unit · not stored in the workbook" footnote), Word `tracked-change` shows the redline as struck old text + inserted new text. An applied write can **drill into its provenance** via a "Show provenance" toggle (`aria-expanded`/`aria-controls`) rendering `ProvenanceDetail`. | `applyProposal`                                                                          |
| error banner            | A `role=alert` panel-error (in `App`).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           | —                                                                                        |
| `Composer`              | The ask box: keyboard-submit input with a label, a grounded/agentic mode toggle, and a send button that flips to Cancel while busy.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              | `send`, `runCommands`, `cancel`                                                          |

Approvals are **fail-closed in the wiring**: nothing actuates until the user clicks Approve, and the
verbatim command line shown is the same `ActuationRequest` that executes — no render-benign /
execute-malicious divergence. The view never reaches the host directly.

## Design tokens — the Starlight / Ramsian language

`src/taskpane/styles.css` is one token system in `:root`; the reference is the Starlight component
spec (Dieter Rams principles). The rules of the language, in order of importance:

1. **Ground is warm paper** (`--paper #fcf9f8`); raised surfaces are plain white (`--surface`) —
   never translucent, blurred, or shadowed. Structure is drawn with **0.5px hairlines**
   (`--hairline`, ink at 40%) and 1–2px ink rules (`--rule`), not tinted boxes.
2. **One functional blue** (`--blue #0057b8`, hover `#00408b`, active `#00336f`) marks the single
   primary action per view, links, and focus. **Red** (`--red #bc000c`) means stop or destroy —
   never emphasis. Everything else is ink (`--ink #1b1c1c`, `--ink-2`, `--ink-3`) on paper.
3. **Status is a lamp plus a word**, never color alone: `--lamp-on` (blue), `--lamp-hold` (gray
   `#c2c6d4`), `--lamp-err` (red). See `.surface-state`, `.status-line`, `.chip .dot`, `.cat`.
4. **Type**: `--sans` (Hanken Grotesk; loaded in `taskpane.html`/`preview.html`, system fallbacks)
   for prose; `--mono` (JetBrains Mono) for data — commands, counts, hashes, provenance. Numerals
   tabular. Micro-labels (`.eyebrow`) are 11px and lowercase.
5. **States shift tone only** — no size change, no shadow, no movement. The single animation is
   the `lampPulse` (busy/streaming), disabled under `prefers-reduced-motion`. Collapse marks are
   − / + set in mono; no rotating chevrons.
6. **Radii**: `--r` (4px) for cards/buttons, `--r-chip` (2px) for chips/tags; knobs (send, stop,
   settings) are circles. Verbatim commands (`.cmd`, `.md-code`) render on the ink plate
   (`--plate` / `--plate-ink`). Inputs are rules, not boxes: bottom hairline, blue only on focus.

The per-surface `data-surface` attribute remains (behavioral hooks + tests), but the accent no
longer forks per host — `--host`/`--link` resolve to the one blue on every surface; the surface
identity is carried by the ink glyph plate and mono labels. Floating layers (popovers, palettes)
use a solid surface with a 1px ink border instead of elevation. Class names
(`panel/ph/pht/pn/pss/unit/chip/card/btn/…`) are unchanged from the feature-wave components.

## Guardrails

The smoke test `src/taskpane/app-render.test.ts` (jsdom, 17 cases) mounts the real `<App/>` over the
`src/taskpane/preview-fixtures.ts` fixtures and asserts every card renders — including the verbatim
plan/write command lines. The fixtures are shared with the preview harness, so the preview and the
test guard the same UI surface. The six newest cases cover: the **skills surface** (signature, `def`
confirmation, prefilled param input, invoke action), an **expanded plan effect** (its target +
dry-run before→after preview), the **citation source-detail popover**, the **Excel formula-first +
entity-card** proposal body, the **Word redline** proposal body, and the **provenance drill-down**.
