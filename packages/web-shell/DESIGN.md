# The Pane Standard

**The design language for the task pane.**
_Version 2.0 · October 2026 · supersedes the Starlight notes in `styles.css`_

---

## What this is

The task pane is an enterprise tool inside Word, Excel, PowerPoint, Outlook, OneNote and Teams.
It sits beside the user's document in a pane that is 320–480px wide. It must look like it came
from one team, every time something changes.

This standard is principles plus tokens. The principles say what we believe. The tokens and
recipes say what to type. A brand pack (`brands/<id>/theme.css`) may remap the token values; the
names and the rules here stay the same.

**Any UI change follows this file.** Agents: read § 9 (the Prompt Block) first, then the recipe
for the part you touch. `bun run lint` runs `scripts/check-design.ts`, which fails on new drift.

Where the values live:

| File | Holds | Who changes it |
| --- | --- | --- |
| `src/taskpane/styles.css` `:root` | Token names and upstream defaults | This standard |
| `brands/<id>/theme.css` | A brand's values mapped onto the token names | The brand pack — do not edit for a UI change |
| `src/taskpane/styles.css` § buttons, § fields | The button and field recipes | This standard |

---

## 1. Product essence (one paragraph)

The pane is a calm, dense, trustworthy work tool, in the tradition of Microsoft's own Office
panes and the best enterprise SaaS (Linear, Stripe Dashboard). It never competes with the
document beside it. One brand red marks the one action that matters. Everything else is ink on
white, with hairline structure. A banker who opens it for the first time sees a product, not a
prototype: every button the same height, every label the same size, nothing that moves when it
should not.

---

## 2. The five principles

### 2.1 One system, no one-offs

Every button is a `.btn`. Every form field is a `.field`. Every size is a token. If a part does
not fit a recipe, the recipe grows first (in this file), then the code uses it. A component class
may set layout (margin, order, width) but never the look.

### 2.2 Density with air

The pane is narrow, so text is small (11–13px) and spacing is tight (4–20px). Air comes from
consistent spacing, not from big type. Body text is 12px. Nothing falls back to the browser's
16px.

### 2.3 One accent, used for the one action

The accent (`--blue`) marks the primary action, the active sidebar item, focus, and the user's
own message. One filled accent button per view. If two things are filled with the accent, one of
them is wrong.

### 2.4 Danger looks different from primary

A brand accent can be red, so a destructive action can never be a filled red button. Danger is an
outlined red button. Destructive copy names what it destroys ("Delete 3 rows").

### 2.5 Equal boxes line up

Controls side by side share one box height: 32px (default) or 24px (small). Every button keeps a
1px border, transparent when unseen, so filled and outlined buttons are the same height. A field
is 32px, the same as a button. Text inside a clipping box gets a line-height that fits its
descenders.

---

## 3. Visual tokens

Use the token. Never type the value. Default values shown; a brand's `theme.css` may override them.

### 3.1 Colour

```
--ink          #1b1c1c   Body text, primary
--ink-2        #424752   Secondary text, labels
--ink-3        #727784   Captions, metadata, disabled text
--paper        #fcf9f8   Pane background
--surface      #ffffff   Raised surfaces: cards, fields, secondary buttons
--hover        #f6f3f2   Hover fill, sidebar background, disabled field
--press        #f0eded   Pressed fill
--press-2      #e4e2e1   Disabled button fill
--hairline     rgba(27,28,28,.4)   Field and button borders
--hairline-soft rgba(27,28,28,.22) Dividers, card borders

--blue         #0057b8   THE accent (a brand may remap it)
--blue-hover   #00408b   Accent hover
--blue-active  #00336f   Accent pressed
--blue-ink     #00408b   Accent as text (links, active labels) — AA on white
--on-accent    #ffffff   Text on a filled accent
--primary-tint (brand)   Selected / active fill (sidebar item, chip); falls back to a blue tint

--red          #bc000c   Danger text and border
--red-tint     #ffdad6   Danger hover fill
--warn-tint    #fff8eb   Caution card fill (interrupted, failed request)
--warn-border  #f5d9a8   Caution card border
```

### 3.2 Typography

```
Sans    --sans   Hanken Grotesk          all UI text (a brand may set its own)
Mono    --mono   JetBrains Mono          commands, counts, ids, provenance
Serif   --serif  Newsreader              quoted document text and redlines only

--fz-xs   10px   badges, micro-hints
--fz-sm   11px   labels, small buttons, captions, card body in compact cards
--fz-md   12px   body text — the pane's base size (.panel)
--fz-lg   13px   buttons, emphasised body
--fz-xl   14px   panel titles, card titles
--fz-icon 16px   a text glyph used as an icon (☆ ⋯ ×)

Weights: 400 body · 500 buttons · 600 labels and card titles · 700 panel titles. Nothing else.
Headings in the empty state may be larger; they are the only exception and live in workspace.css.
```

### 3.3 Spacing scale

```
--s1 4px · --s2 6px · --s3 8px · --s4 12px · --s5 16px · --s6 20px
```

Pane gutter is `--s5`. Gap between stacked sections is `--s6`. Gap inside a group is `--s4`.
Gap between a label and its control is `--s2`. Gap between buttons in a row is `--s2`.

### 3.4 Radius

```
--r        4px    buttons, fields, cards, panels
--r-chip   2px    tags
9999px            round knobs only (avatar, send button upstream)
```

A brand pack may give the chat a rounder look (message bubbles, composer, send button); that is
the only place other radii appear.

### 3.5 Lines and elevation

Structure is drawn with 1px `--hairline-soft` dividers and borders, not with shadows. Add no new
shadows; the composer's focus ring in the brand pack is the one deliberate exception. No
gradients.

### 3.6 Icons

Monoline SVG, 24×24 viewBox, 1.75px stroke, no fill (`.tw-glyph`). Size by width/height:
18px in the sidebar and header, 16px in a `.btn`, 14px in a `.btn.sm`. An icon-only control has
an `aria-label`. An icon never carries meaning that its label does not.

---

## 4. Buttons — the one button system

Every button is `.btn` plus modifiers. Defined in `styles.css` § buttons.

| Recipe | Class | Height | Text | Use |
| --- | --- | --- | --- | --- |
| Primary | `.btn.pr` | 32 | 13px/500, white on red | The one main action of a view: Apply, Send, Accept change |
| Secondary | `.btn` | 32 | 13px/500, ink on white, 1px `--hairline` border | Every other action: Cancel, Reject, Refresh |
| Quiet | `.btn.quiet` | 32 | 13px/500, `--ink-2`, no border; hover fills `--hover` | Low-weight actions in a busy row: Skip, Show more |
| Danger | `.btn.danger` | 32 | 13px/500, red text, red border, white fill | Destroys something: Delete, Clear history |
| Icon | `.btn.icon` | 32×32 | — | Close, more; needs `aria-label` |
| Small | add `.sm` | 24 | 11px/500 | In cards and compact rows: Finish it, Dismiss, Undo |
| Text link | `.link` | auto | `--blue-ink` | Inline in a sentence |

Rules:

1. One `.btn.pr` per view. A card with its own decision (approval, resume) counts as a view.
2. Order in a row: primary first, then secondary, then quiet. Danger stands alone or last.
3. A row of buttons uses one size. Never mix `.sm` and default in one row.
4. Labels are verbs, sentence case, 1–3 words: "Accept change", not "OK" or "ACCEPT".
5. Disabled is `disabled`, not a colour. Say why nearby when it is not obvious.
6. A component class on a button (e.g. `error-card-retry`) may only set layout.

Named recipes that predate `.btn` and are allowed by the check: `tw-rail-item` (sidebar),
`tw-rail-toggle`, `tw-head-btn` (header), `tw-modal-close`, `cite` (citation chip), `suggestion`
and `workspace-action` (start cards), `palette-item` (the `/` menu). A new kind of button is
added here first, then to `RECIPES` in `scripts/check-design.ts`.

---

## 5. Pattern library

### 5.1 Field

```html
<label class="field">
  <span class="field-label">Format</span>
  <select class="field-control">…</select>
</label>
```

Label 11px/600 `--ink-2`, 6px above a 32px control with a 1px `--hairline` border and `--r`.
Inputs, selects and textareas inherit the pane font; they never fall back to 16px.

### 5.2 Section label (eyebrow)

`.eyebrow`: 11px/500, `--ink-2`, lowercase. One per section, above the section's first item.
Do not use it as a heading for a single control.

### 5.3 Panel (sidebar modal)

Title 14px/700 in the header beside a `.btn.icon`-sized close. Body padding `--s5`. Sections
stack with `--s6` between them (`.settings-stack`).

### 5.4 Caution card

`.error-card`: `--warn-tint` fill, `--warn-border`, `--r`. Title 11px/600, body 11px `--ink-2`,
actions are `.btn.sm` in a `.result-card-actions` row with `--s3` above.

### 5.5 Sidebar

60px rail on `--hover`. Items are 58×48, glyph 18px over an 11px/600 label with a 14px line
height. Active item: `--primary-tint` fill, `--blue-ink` text. Settings sits at the bottom.

---

## 6. What we never do

- Never type a colour, size or radius that has a token. Use the token.
- Never leave text without a size inside a new container; the base is 12px from `.panel`.
- Never style a `<button>` from scratch. Use `.btn` and modifiers.
- Never put two filled red buttons in one view.
- Never make danger a filled red button.
- Never put a 24px and a 32px control side by side.
- Never clip text with `overflow: hidden` and a line-height smaller than the glyphs need.
- Never use shadows or gradients for structure.
- Never use emoji as icons; text glyphs (☆ ⋯ ×) use `--fz-icon`.
- Never change `brands/<id>/` files for a UI change.
- Never ship a UI change without screenshots at 320, 360 and 480px.

---

## 7. States and pane sizes

Hover shifts fill only (`--hover`, `--press`). Focus is the 2px `--ring` outline from the global
`:focus-visible`. Disabled controls stay in place and keep their size. During a turn or approval,
task and context controls disable; navigation and inspection stay available.

Check every change in the preview (`bun run --filter @ge/web-shell preview`, then
`?surface=<host>&demo=1`) at 320, 360 and 480px wide and 600 and 760px tall. Chips and button
rows wrap. Long content scrolls inside its region. The composer stays outside the thread's
scroll. Honour reduced motion.

---

## 8. Enforcement

`bun run lint` ends with `bun scripts/check-design.ts`. It fails when a file gains:

- a `font-size` that is not a `--fz-*` token,
- a `border-radius` that is not `--r`, `--r-chip`, `0`, `50%` or `9999px`,
- a raw hex / `rgb()` / `hsl()` colour outside a token declaration,
- a `<button>` that is not a recipe from § 4.

Existing code is held at `packages/web-shell/design-baseline.json`. When you clean a file up, run
`bun run design:check --update` so the lower count becomes the new limit.

---

## 9. The Prompt Block

Paste or read this before any UI change to the pane.

```
You are changing the Contoso-AI task pane (packages/web-shell). Follow packages/web-shell/DESIGN.md.

Tokens only: colours var(--ink|--ink-2|--ink-3|--surface|--hover|--press|--hairline|
--hairline-soft|--blue|--blue-hover|--blue-ink|--on-accent|--primary-tint|--red|--red-tint|
--warn-tint|--warn-border); sizes var(--fz-xs 10|--fz-sm 11|--fz-md 12 base|--fz-lg 13|
--fz-xl 14|--fz-icon 16); spacing var(--s1 4|--s2 6|--s3 8|--s4 12|--s5 16|--s6 20);
radius var(--r 4) or var(--r-chip 2). Weights 400/500/600 (700 panel titles only).

Buttons: always class "btn" + modifiers — pr (one per view), quiet, danger (outlined, never
filled), icon (aria-label), sm (24px; default is 32px). One size per row. Verb labels.
Fields: label.field > span.field-label + .field-control (32px). Section label: .eyebrow.
A component class on a control sets layout only.

Before you finish: bun run typecheck, bun run test, bun run lint (includes the design check),
then screenshot the preview at 320/360/480px and compare against this file. Report what you
checked.
```

---

## 10. Versioning

| Version | Date | Change |
| --- | --- | --- |
| 2.0 | 2026-10-16 | Principles + tokens shape, one button system, field recipe, design check |
| 1.0 | — | Starlight notes (blue theme, top toolbar) |
