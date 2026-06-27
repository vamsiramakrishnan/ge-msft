# Task Pane Design System

## Direction

The pane uses a restrained Office-native product language. It should feel closer to a command center embedded in Word or Excel than to a standalone AI chatbot.

## Visual Hierarchy

1. Host identity, readiness, and the current surface.
2. Three primary actions for the active host.
3. Grounding scope and routing state as compact rows.
4. Conversation, citations, command traces, and gated approvals.
5. Secondary action catalog and session skills behind hover/focus disclosure.

## Chrome

Use flat surfaces, hairline borders, and restrained elevation only for transient popovers or pinned approval rails. Avoid nested cards, large gradients, decorative symbols, and persistent secondary panels.

## Color

The Microsoft host accent is the main interaction color. Gemini color is reserved for the small product identity mark. Semantic states use stable green, amber, and red tokens.

## Progressive Disclosure

Secondary information appears through hover, focus, or an explicit edit action:

- Context chips expand from a single grounding row.
- Catalog routing expands only when editing skills/connectors.
- Skills open as a compact flyout.
- Quick actions remain collapsed until the user explores the drawer.
- Citations and provenance expose detail on demand.

## Interaction

Every hover affordance must also open through keyboard focus. Primary actions remain visible because they are task accelerators; secondary action lists remain discoverable but not visually dominant.

## Responsive Behavior

The task pane is designed from the narrow Office pane outward. At narrow widths, controls stack into single-column rows, popovers become scrollable sheets inside the pane, and dense metadata collapses before primary actions do. At wider preview widths, primary actions can return to multi-column layout while the conversation remains readable. At short heights, header and action chrome compress so approvals and the composer stay reachable.
