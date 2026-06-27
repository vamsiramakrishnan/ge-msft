---
target: taskpane
total_score: 30
p0_count: 0
p1_count: 1
timestamp: 2026-06-27T08-57-12Z
slug: src-taskpane-components-app-tsx
---

# Critique — Gemini Enterprise M365 Task Pane (`packages/web-shell`)

Target: the task pane (`src/taskpane/components/App.tsx` + components + `styles.css`).
Register: product. Reviewed from source + computed CSS; browser automation unavailable
(no dev server, Office-iframed pane). Deterministic detector: clean (0 findings) on TSX markup.

## Design Health Score (Nielsen's 10)

| #         | Heuristic                   | Score     | Key issue                                                                                                                                                                                     |
| --------- | --------------------------- | --------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1         | Visibility of system status | 3         | Strong (readiness pill + aria-busy + run-steps), but status is split: two busy signals (`busy` vs `actionBlocked`) and the pill says "Decision needed" while the gate rail is the real locus. |
| 2         | Match system / real world   | 3         | "research unit grounding scope" / "capability closure" leak internal vocabulary into labels + catalog tooltip.                                                                                |
| 3         | User control & freedom      | 3         | Reject/Cancel everywhere, but no in-pane undo after Accept — reversibility is asserted in copy, lives in the host.                                                                            |
| 4         | Consistency & standards     | 3         | One token system, shared input vocabulary — but three lexicons for the same three output modes (see P1).                                                                                      |
| 5         | Error prevention            | 4         | Excellent: fail-closed `{{}}` guard, planner-confirm front door, dry-run before gate, inert non-http hrefs. The product's spine.                                                              |
| 6         | Recognition over recall     | 3         | `/`-palette + `@`-mentions discoverable, but action meaning is hover-gated; a returning user must hover to recall what a verb does.                                                           |
| 7         | Flexibility & efficiency    | 3         | Power path (slash/scope/mentions) + chips coexist, but primary actions duplicated between command center and quick-action drawer.                                                             |
| 8         | Aesthetic & minimalist      | 2         | Violates its own "quiet by default": 5–6 equal-weight sections render above the first message.                                                                                                |
| 9         | Error recovery              | 3         | Bootstrap fatal states exemplary; in-pane errors are bare strings with `⚠` and degraded/blocked status lines offer no next action.                                                            |
| 10        | Help & documentation        | 3         | Inline entrypoints popover + catalog `i` tooltip are nice, but all hover-gated — easy to never find.                                                                                          |
| **Total** |                             | **30/40** | **Good** — strong on safety/prevention, weak on composition + lexical consistency.                                                                                                            |

## Anti-Patterns Verdict

**LLM assessment:** Largely escaped the AI-card look — reads as a deliberate, hand-tuned product.
Cards genuinely flattened (base `box-shadow: var(--sh)` → override `box-shadow:none`); eyebrow tell
removed (`.cat` uppercase/tracking zeroed); host-accent theming is real (per-`data-surface` `--host`
propagates through rings/buttons/bubbles/hover tints). Residual tells: single-char Unicode glyphs as
the icon language (`→` send, `◼` cancel, `?/+/>` actions), and section sameness (5–6 identical
hairline 8px-radius rectangles stack before any content, so nothing leads).

**Deterministic scan:** `detect.mjs` clean — 0 findings on the TSX markup. No gradient-text,
eyebrow, side-stripe, or flat-palette hits. Agrees with the LLM read that the slop families are absent.

**Visual overlays:** none — browser injection unavailable in this environment (no running dev server;
the pane is normally Office-iframed). Fallback: source + computed-CSS review only.

## Overall Impression

A confident, safety-first command surface whose approval gate is best-in-class for an AI write tool —
and whose first-run composition undercuts it by rendering an admin console's worth of chrome before
the user can ask a question. The thesis ("every write is earned") is executed with conviction; the
"quiet by default" principle is not. Biggest opportunity: give the thread + composer visual primacy
on open and collapse catalog/skills to single-line affordances.

## What's Working

1. **The approval gate is the standout.** Verbatim command rendering (the same request that executes),
   per-effect expandable dry-run with before→after diffs, and reassurance copy at the moment of
   mutation ("will not run until you approve … a reversible, provenanced change"). Makes the thesis
   visible, not just claimed.
2. **Host-native theming is real, not cosmetic.** `data-surface` swaps `--host` and it propagates
   through focus rings, glyph, primary buttons, and message bubbles — switch Word→Excel and the pane
   actually becomes green throughout. Delivers "host first."
3. **Untrusted-content hygiene shows up in the UI layer.** `safeHttpUri` renders `javascript:`/`data:`
   URIs as inert text in citations + provenance; the markdown renderer never uses dangerous HTML.
   A design that treats source content as hostile, visibly.

## Priority Issues

**[P1] First-run renders too much chrome — the pane fights its own "quiet by default" principle.**

- Why: 5–6 equal-weight sections (command center, context tray, catalog, suggestions, skills) stack
  above the thread on mount; catalog auto-loads. Cognitive-load checklist fails ~4 of 8. A first-timer
  meets an admin console; the warm empty-state copy is the _last_ thing they see, below the fold.
- Fix: collapse `catalog` + `skills` to one-line affordances by default (the flyout machinery already
  exists — make the panel itself the trigger), demote the catalog editor behind a settings entry, give
  thread/composer load-time priority.
- Command: `/impeccable quieter` then `/impeccable layout`.

**[P2] Three inconsistent names for the same three output modes.**

- Why: output kind chat/annotation/write is labeled Ask/Review/Change (SurfaceCommandCenter shortMode),
  Answer/Review gate/Preview gate (actionMode), Change/Review/Ask + `?/+/>` icons (QuickActionBar), and
  Ask/Preview comments/Preview write (param-form CTA). The user must learn Change = write = Preview
  write = `>`; recognition collapses and it reads machine-generated.
- Fix: one canonical lexicon + one icon per mode defined in `contracts` next to `output`, consumed
  everywhere.
- Command: `/impeccable clarify`.

**[P2] Hover-disclosed content lacks a discoverable keyboard affordance.**

- Why: `.unit`/`.catalog`/`.skills`/`.action-drawer` reveal via `:hover`/`:focus-within` (so it's not
  strictly hover-only), but there's no visible chevron/`aria-expanded` telling a keyboard user the
  reveal exists; `ContextTray`'s `.unit-peek` is `aria-hidden`. PRODUCT.md commits to "focus path for
  every hover-disclosed control" — the mechanism is there, the discoverability isn't.
- Fix: add a persistent disclosure control (chevron + `aria-expanded`) to each collapsible section.
- Command: `/impeccable clarify` + accessibility `/impeccable audit`.

**[P2] Muted text contrast is AA-compliant but tight; verify margin on `--psoft`.**

- Why: recomputed WCAG ratios — `--psoft #74716c` ≈ 4.66:1 on `--bg`, dropping toward ~4.6:1 on
  `--fill-2` and used at 10.5–11.5px (`.muted`, `.pss`, `.s-detail`, `.card .w` reassurance copy).
  Passes, but with almost no margin, and it's the trust-critical reassurance copy. `--muted` (~5.3:1)
  and `--soft` (~5.7:1) are comfortable.
- Fix: nudge `--psoft` ~one step darker (e.g. ~`#6a675f`) for body-size use, or reserve it for ≥14px;
  keep `--soft`/`--muted` for small text.
- Command: `/impeccable audit` (contrast pass).

**[P3] Iconography is single-char Unicode glyphs; send/cancel read as placeholders.**

- Why: `→` send / `◼` cancel (ambiguous), `?/+/>` action icons, `◆` entity. Most visible "AI made
  this" residue; undercuts the otherwise enterprise feel.
- Fix: adopt a minimal consistent set (Fluent UI icons are host-native here) for at least
  send/cancel/attach/detach and the three modes.
- Command: `/impeccable polish`.

## Persona Red Flags

**Jordan (first-timer) — most at risk.** Lands on auto-loaded catalog + skills + suggestions + a 3-up
action grid before the warm empty-state copy (which sits below all of it). `/`, `@`, scope radiogroup,
and the "+N" drawer are powerful but unexplained at rest.

**Sam (accessibility / keyboard / SR).** `.unit-peek` is `aria-hidden` and chip reveal depends on
focusing into a section with no visible "expandable" cue. The Composer palette is `role="listbox"`
with `role="option"` children but no roving focus / arrow-key nav — listbox semantics, list-of-buttons
behavior. Glyph-only send/cancel rely on `aria-label` with no visible text fallback.

**Riley (stress tester).** Two busy concepts (`state.busy` vs `actionBlocked`) disable different
regions during a gate → inconsistent disabled states under fast clicking. Hand-rolled markdown
renderer is a fuzzing target for partial/streamed tables + unclosed code fences.

## Minor Observations

- Dead first `:root` palette (`--grad` blue→pink, `--brand #8b5cf6`) is overridden but `--grad` is
  _not_ redefined, so `.snd`/`.btn.pr`/`.fatal-action`/`.card-top` base still use the old gradient —
  bootstrap chrome diverges from the host-accent app.
- `KIND_DOT` in ContextTray uses hardcoded hex instead of tokens (off-palette dots).
- `.cat::before` dot duplicates the status-line dot (two dots per card).
- SurfaceCommandCenter returns `null` when `actions.length === 0` — a capability-restricted surface
  silently loses its header with no explanation (no empty state).
- Three `aria-live="polite"` regions (thread log, run-steps, approval cards) can over-announce during
  a streaming + gating turn.

## Questions to Consider

1. If "every write is earned," why is undo/rollback invisible in the pane _after_ apply? The gate
   reassures before the leap and goes quiet after.
2. Does a knowledge worker mid-document ever pick a planner skill + command skill + connector data
   stores from dropdowns? Is the catalog admin config wearing a task-pane costume?
3. You flattened the cards beautifully but left six equal-weight sections. If everything is quiet,
   nothing leads. What is the one thing this pane wants the user to do on open?
4. The model emits one command grammar — why does the human-facing surface speak three dialects of it?
