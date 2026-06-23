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
- **Cards** — independent toggles for each state slice (Context, Thread, Suggestions, Run steps,
  Plan, Write, Proposals, Error, Busy), plus **All on** / **Idle / empty** presets.

The preview is view-only and never ships in the add-in bundle (`vite.preview.config.ts` is separate
from `vite.config.ts`).

## Component map

The panel hierarchy, top to bottom (`src/taskpane/components/`):

| Component | Renders | Controller methods |
| --- | --- | --- |
| `App` | Panel shell: `data-surface`, `aria-busy`, header (agent identity), and the `<main>` thread region landmark. Wires every child to the controller. | `refreshContext`, `attach`, `detach` (via `onToggle`), `onAutomate`, `dismissSuggestion` |
| `ContextTray` | The research-unit grounding scope as chips, split into "Attached sources" / "Available to attach" `role=list` groups; each chip is an attach/detach `<button>`. | `refreshContext`, `attach`, `detach` |
| suggestions `<section>` | Clickable suggestion cards (in `App`). | `onAutomate`, `dismissSuggestion` |
| `MessageThread` | The grounded conversation (`role=log`, `aria-live=polite`): user/assistant bubbles, streaming caret, a "Sources" eyebrow + citation pills (http(s)-only hrefs). | — (read-only view of `messages`) |
| `RunSteps` | The command-loop transcript: an "Activity" section with an ordered, polite live region of run steps. | — (read-only view of `steps`) |
| `PlanApprovalCard` | Fail-closed plan-level gate (ADR-0005): a `role=region`/`aria-live` card listing each effect's **verbatim** command line via `renderCommandLine`, with Approve / Reject. | `approvePlan`, `rejectPlan` |
| `WriteApprovalCard` | Fail-closed per-write gate (ADR-0004): a `role=region`/`aria-live` card showing the **verbatim** `pending.command`, with Approve / Reject. | `approvePendingWrite`, `rejectPendingWrite` |
| `ProposalCard` | The reversible-write review: one card per staged proposal in its status; pending exposes an "Accept change" `<button>`. | `applyProposal` |
| error banner | A `role=alert` panel-error (in `App`). | — |
| `Composer` | The ask box: keyboard-submit input with a label, a grounded/agentic mode toggle, and a send button that flips to Cancel while busy. | `send`, `runCommands`, `cancel` |

Approvals are **fail-closed in the wiring**: nothing actuates until the user clicks Approve, and the
verbatim command line shown is the same `ActuationRequest` that executes — no render-benign /
execute-malicious divergence. The view never reaches the host directly.

## Design tokens

`src/taskpane/styles.css` is one token system in `:root` so every card shares the same chrome. The
palette is lifted from `docs/mockups/*.html`. Token groups:

- **brand** — `--grad` (the Gemini gradient), `--brand` (fill `#8b5cf6`), `--brand-ink` (the
  AA-compliant variant for small brand text labels), `--brand-soft`, `--brand-tint`.
- **host accent** — `--host`, `--link` (Word default; overridden per surface).
- **ink + surfaces** — `--ink`, `--ink-2`, `--soft`, `--psoft`, `--muted`, `--bg`, `--fill`, `--fill-2`.
- **lines** — `--pl`, `--line`.
- **status** — `--ok`, `--warn`, `--danger` (+ `--danger-bg` / `--danger-line`), `--teal`.
- **spacing** — `--s1`…`--s6`.
- **radii** — `--r-sm`, `--r-md`, `--r-lg`, `--r-pill`.
- **typography** — `--pf`, `--mono`, `--fz-xs`…`--fz-xl`.
- **elevation** — `--sh`, `--ring`.

Per-surface accent override: `.panel[data-surface='…']` re-points `--host` and `--link` to the host's
color (Word `#185abd`, Excel `#107c41`, PowerPoint `#c43e1c`, Outlook `#0f6cbd`, Teams `#5b5fc7`)
while the brand gradient stays constant. Existing class names (`panel/ph/pht/pn/pss/unit/chip/card/btn`)
are unchanged from the mockups.

## Guardrails

The smoke test `src/taskpane/app-render.test.ts` (jsdom, 11 cases) mounts the real `<App/>` over the
`src/taskpane/preview-fixtures.ts` fixtures and asserts every card renders — including the verbatim
plan/write command lines. The fixtures are shared with the preview harness, so the preview and the
test guard the same UI surface.
