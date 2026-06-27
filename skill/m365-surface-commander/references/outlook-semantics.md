# Outlook semantics

Load this when the active surface is Outlook. Cross-surface table:
[capability-map.md](capability-map.md).

**Reading:** `read` returns the open mail item (or the compose draft). **Nothing auto-sends** — every
Plane-A write mutates the open **compose draft**, and the draft only leaves on the user's Send, which
passes through the on-send veto gate. So every compose write is inherently draft-reviewable.

**Core verbs:** `mail "body"` stages a reviewable reply; `compose "Subject" "body"` opens a new draft
(recipients left for the user — never auto-addressed).

**Specialized (`/`):** `/add-attachment` (attach the working doc / a generated file),
`/set-recipients` (to/cc/bcc), `/set-subject`, `/set-body`, `/prepend-body`, `/add-categories`,
`/set-sensitivity-label`, `/set-internet-headers` (custom `x-*` only — a provenance carrier),
`/compose-appointment` (open a meeting form), `/save-draft`.

**Gotchas:** **importance and follow-up flags are Graph-only** (not in the compose API) — setting them
is a Plane-B estate write that needs estate auth, not a `/`-on-the-draft. `/add-categories` on a
_received_ item commits **immediately** (no on-send safety net) and the category must already exist in
the master list. The URI for `/add-attachment` is untrusted — it must be allowlisted. Server-side
`sendMail` is intentionally **not** modeled (it bypasses the on-send gate); the reviewable path is
`mail`/`compose` + the gate.
