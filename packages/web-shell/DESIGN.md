# Task pane design system

The shipped pane uses a warm paper surface, near-black ink, and one functional blue. These tokens
are defined in `src/taskpane/styles.css`; workspace extensions live in `workspace.css`.

## Tokens

| Purpose | Token | Value |
| --- | --- | --- |
| Main text | `--ink` | `#1b1c1c` |
| Secondary text | `--ink-2` | `#424752` |
| Paper | `--paper` | `#fcf9f8` |
| Raised surface | `--surface` | `#ffffff` |
| Interaction | `--blue` | `#0057b8` |
| Link text | `--blue-ink` | `#00408b` |
| Error | `--red` | `#bc000c` |
| Selected source | — | `#eff5fc` with `#bacce0` border |

Use the shared spacing scale: 4, 6, 8, 12, 16, and 20 px. The base radius is 4 px. Source and
filter chips use a pill outline. Prose uses Hanken Grotesk with system fallbacks; commands use
JetBrains Mono; document redlines use the existing serif stack. Font files are not required for
controls to remain usable.

## Hierarchy

The header identifies the product and host. Navigation is compact. Active context stays visible.
The empty conversation starts with three relevant tasks; a populated conversation gives space to
the answer. Search and specialist configuration use the existing dialog. Approval is a separate
scrollable region above the composer.

Do not show implementation architecture in the footer. Do not describe every mutation as reversible
when the host may not support that guarantee. Show exact commands in review, where they help the
user understand the operation being approved.

## Interaction states

Use the functional blue for selected controls and focus. Pair state with text. Disable task and
context mutation during a turn or approval, while leaving navigation and inspection available.
The composer offers a stop control while working. Streaming artifacts cannot be inserted.

## Pane sizes

Inspect 320, 360, and 480 px widths and 480, 600, and 760 px heights in the preview harness.
Chips and filters wrap. Long content scrolls inside its region. Keep the composer outside the
conversation's scrolling area. Avoid hover-only access and honor reduced-motion preferences.
