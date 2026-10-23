---
name: ui-change
description: Use for ANY change to how the task pane looks or is laid out — a button, label, panel, card, font size, spacing, alignment, colour, icon, sidebar, or "move X to Y" in packages/web-shell. Applies the Pane Standard (packages/web-shell/DESIGN.md) and verifies with screenshots at 320/360/480px before reporting done.
---

The pane has one design language: `packages/web-shell/DESIGN.md` (the Pane Standard).
Every UI change follows it. Do not invent sizes, colours or button styles.

## Before you edit

1. Read `packages/web-shell/DESIGN.md` § 9 (Prompt Block) and the recipe for each part you touch
   (§ 4 buttons, § 5 fields / panels / cards / sidebar).
2. Find the existing part in the code. Reuse its recipe. If nothing fits, extend the recipe in
   `DESIGN.md` and `styles.css` first, then use it. Say so in your report.
3. Check the brand: a brand's values come from `brands/<id>/theme.css`. Do not edit that file for
   a UI change.

## While you edit

- Buttons: `className="btn"` plus `pr` / `quiet` / `danger` / `icon` / `sm`. One `pr` per view.
  One size per row. Danger is outlined.
- Fields: `label.field > span.field-label + .field-control`.
- Sizes, spacing, radius and colours: `var(--fz-*)`, `var(--s*)`, `var(--r)`, colour tokens.
- A component class sets layout only (margin, order, width).
- Text inside a clipping box gets a line-height that fits descenders.

## Before you report done

1. `bun run typecheck`, `bun run test`, `bun run lint` (lint includes `scripts/check-design.ts`).
   A design-check failure means new drift: fix it with tokens or recipes, never by raising the
   baseline. Run `bun run design:check --update` only after a clean-up lowers a count.
2. Start the preview for THIS checkout (`bun run --filter @ge/web-shell preview`; if port 17100
   is taken by another checkout, run `bunx vite --config vite.preview.config.ts --port 17101
   --strictPort` from `packages/web-shell`).
3. Screenshot every changed state at the three widths:
   `node scripts/pane-shots.mjs --url http://localhost:<port> --surface <host> [--panel <name>]
   [--query resume=1] --out <scratchpad>/shots`
4. Look at each screenshot. Check against DESIGN.md § 6 (What we never do): equal heights in a
   row, no 16px text, one filled red button, labels not clipped, nothing overflows at 320px.
5. Report: what changed, the screenshots you checked, and any older drift you saw nearby but did
   not touch (offer it as a separate clean-up).
