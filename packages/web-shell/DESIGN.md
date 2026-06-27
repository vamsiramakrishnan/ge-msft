# Task Pane Design System

Captured from the shipped `src/taskpane/styles.css`. The stylesheet layers a later
"product polish" `:root` override on top of the original mockup palette; the values below are
the **effective** ones (the override wins in the cascade). Keep this doc in lockstep with that
override block — if you retune tokens there, retune them here.

## Direction

A restrained, Office-native command surface. The pane reads as a command center embedded in
Word or Excel, not a standalone AI chatbot: flat surfaces, hairline borders, compact rows over
cards, and the host's own accent as the interaction color. Chrome is quiet by default; depth and
detail appear only on hover, focus, edit, or when a decision is required.

## Theme

Light, warm-neutral. A near-white warm-grey body (`#fbfaf8`) with pure-white raised panels
(`#ffffff`), near-black ink, and a single Material-purple product accent that mostly yields to
the per-surface host color. This is an existing, committed identity — preserve it; do not
re-tint or re-theme on a whim.

## Color

Tokens are CSS custom properties on `:root`, with the host accent overridden per surface on
`.panel[data-surface=…]`.

### Brand / product mark

- `--brand` `#6750a4` · `--brand-ink` `#56418f` (AA text) · `--brand-soft` `#d9d1ee` · `--brand-tint` `#f8f6fc`
- Gemini identity mark gradient: `linear-gradient(135deg, #305fce, #6750a4)` — reserved for the
  small avatar/mark only, never for surfaces or buttons.

### Host accent (the interaction color — overridden per surface)

- `--host` / `--link`: Word `#185abd` · Excel `#107c41` · PowerPoint `#c43e1c` · Outlook `#0f6cbd` · Teams `#5b5fc7`
- The host accent drives focus rings, links, primary buttons, selected states, and hover tints
  (via `color-mix(in srgb, var(--host) N%, …)`). Brand purple is reserved for product identity.

### Ink + surfaces

- `--ink` `#1f1f1f` · `--ink-2` `#2d2d2d` · `--soft` `#66645f` · `--psoft` `#74716c` · `--muted` `#6b6964`
- `--bg` `#fbfaf8` · `--panel` `#ffffff` · `--fill` `#f5f4f1` · `--fill-2` `#f8f7f4`
- `--pl` `#ebe8e2` (panel hairline) · `--line` `#ddd9d0` (structural hairline)

### Semantic status (stable across surfaces)

- success `--ok` `#0f6f3f` · warning `--warn` `#8a5a00` · danger `--danger` `#b42318`
  (`--danger-bg` `#fff2f0`, `--danger-line` `#f1bbb4`) · citation `--teal` `#087568`
- Status is carried by a colored dot + text and a 3px card top-rail, not by full-bleed fills.

## Typography

- Family: `--pf: 'Aptos', 'Segoe UI', system-ui, sans-serif` (one family carries the whole pane);
  `--mono: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace` for commands, formulas,
  citations, and the `cmd` algebra.
- Fixed px scale (not fluid): `--fz-xs` 10.5 · `--fz-sm` 11.5 · `--fz-md` 12.5 · `--fz-lg` 13.5 ·
  `--fz-xl` 14. Scale steps down at `≤340px` and is mirrored in markdown content sizes.
- Weights run heavy for a UI: labels/titles 700–760, with finer steps (620–760) used to grade
  emphasis in dense rows. Body/prose line-height ~1.5–1.55.
- The "polish" layer **de-emphasizes eyebrows**: `.eyebrow` / `.cat` / `.cites-h` drop
  `text-transform: uppercase` and tracking (letter-spacing 0). No tracked all-caps kickers.

## Spacing

Six-step scale: `--s1` 4 · `--s2` 6 · `--s3` 8 · `--s4` 11 · `--s5` 14 · `--s6` 16 (px).
Section blocks sit at a 14px inline margin (`margin-inline: 14px`), tightening to 10/8px at
narrow widths and centering to `min(720px, …)` at preview width (≥520px).

## Radii

`--r-sm` 6 · `--r-md` 8 · `--r-lg` 10 · `--r-pill` 999 (px). Pill radius is used for chips,
quick actions, citations, and status badges; sm/md for panels, inputs, and cards.

## Elevation

Flat by default. Borders do the work; shadow is reserved for transient/raised surfaces only.

- `--sh` `0 14px 34px rgba(31,31,31,.12), 0 2px 8px rgba(31,31,31,.07)` — popovers, flyouts,
  catalog grid, citation detail.
- `--ring` `0 0 0 3px color-mix(in srgb, var(--host) 18%, transparent)` — focus-within on inputs.
- The pinned gate rail uses an upward shadow (`0 -12px 28px rgba(31,31,31,.08)`) to read as
  docked. Persistent panels (`.surface-center`, `.unit`, `.catalog`, `.skills`) ship `box-shadow: none`.

## Components

Every interactive control carries default / hover / focus-visible / disabled states; hover and
focus-visible are styled together so nothing is hover-only.

- **Header** (`.ph`): product mark + "Gemini Enterprise" + grounding subtitle. Flat, hairline base.
- **Surface command center** (`.surface-center`): host glyph, title, readiness pill
  (`ready` / `busy` / `gate`), and up to three primary `.surface-action` rows (mode badge +
  label + meta). Single-column on narrow panes, 3-up at preview width.
- **Context tray** (`.unit`): one grounding row that expands attach/detach chips on hover/focus.
- **Catalog / Skills** (`.catalog`, `.skills`): compact summary rows that open flyouts (`--sh`).
- **Quick action drawer** (`.quick-actions.action-drawer`): collapsed by default; the secondary
  action list opens upward on hover/focus. Output kind is color-coded (write = green,
  annotation = amber) on icon + meta.
- **Thread** (`.thread`): assistant rows (mark + bubble + citations), user bubbles (host-tinted),
  streamed markdown (`.md-content`), and a blinking caret while streaming.
- **Decision gate rail** (`.gate-rail`): pinned, scrollable region holding the CommandPlan,
  Plan-approval, and Write-approval cards. The single place writes are confirmed.
- **Cards** (`.card`): flat, hairline, 3px top-rail colored by status
  (brand/ok/danger/warn). No nesting.
- **Composer** (`.comp`): scope toggle, optional `/`-palette + `{{param}}` fill form, textarea +
  send/cancel. Focus-within raises the `--ring`.
- Inputs (`select`, text, textarea) share one vocabulary: hairline border, sm radius, host-tinted
  focus ring.

## Motion

State-conveying only; ~160ms `ease` on border/background/transform for action rows and controls.
No orchestrated load sequences, no decorative motion. The streaming caret blink is the one
ambient animation. `@media (prefers-reduced-motion: reduce)` collapses all animation/transition
durations to ~0.

## Responsive Behavior

Designed from the narrow Office pane outward, with height-aware compression — the pane is often
short, not just narrow.

- `≤380px` / `≤340px`: section margins tighten (10/8px), headers and dense rows stack to a single
  column, the type scale steps down, popovers become full-width scroll sheets.
- `≥520px` (browser preview): content centers at `min(720px, calc(100% - 28px))`, primary actions
  return to a 3-up grid, flyouts center via `translateX(-50%)`.
- `≤680px` / `≤560px` tall: header/section padding compresses, subtitles and secondary metadata
  drop, the gate rail caps at `42vh` and scrolls — primary actions and the composer collapse last.

## Accessibility

WCAG AA: body ≥4.5:1, large/bold ≥3:1 (`--brand-ink` is the AA-safe purple for small text; the
`#6750a4` fill token is not used for small text). Visible 2px `:focus-visible` outline in the host
color. Every hover-disclosed control opens on keyboard focus (`:focus-within`). Reduced-motion is
honored. No information is conveyed by color alone — status pairs a dot/badge with text.
