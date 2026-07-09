---
title: Outlook Semantics
kind: reference
skill: m365-surface-commander
surface: outlook
topics: [mail, compose, attachments, staged-draft]
load_when: The active surface is Outlook or a plan mentions mail, thread, attachment, reply, or draft.
---

# Outlook semantics

Load this when the active surface is Outlook. Cross-surface table:
[capability-map.md](capability-map.md).

**Reading:** `read` returns the open mail item (or the compose draft). **Nothing auto-sends** — every
Plane-A write mutates the open **compose draft**, and the draft only leaves on the user's Send, which
passes through the on-send veto gate. So every compose write is inherently draft-reviewable.

**Progressive disclosure:** build slash commands from the live Outlook capability scan for the current
item. Prefer `read` / `list` / `inspect` / `properties` / `open` before any write: inspect whether the
surface is read or compose, list attachments/categories/recipients when exposed, read body coercion and
IDs, and open related drafts/items through the host. If mode or capability is missing, stale, or
ambiguous, fail closed: hide the write command or return a capability error instead of trying a best
effort Office.js or Graph call.

**Core verbs:** `mail "body"` stages a reviewable reply; `compose "Subject" "body"` opens a new draft
(recipients left for the user — never auto-addressed).

**Specialized (`/`) when live-capability advertised:** `/set-body` and `/prepend-body` mutate only the
open compose draft body; `/add-attachment` attaches an allowlisted working doc / generated file to the
open draft; `/set-recipients` changes to/cc/bcc **only on a draft**; `/set-subject`, `/add-categories`,
`/set-sensitivity-label`, `/set-internet-headers` (custom `x-*` only — a provenance carrier),
`/compose-appointment` (open an appointment/meeting draft, never send invitations), `/save-draft`.

**Gotchas:** **importance and follow-up flags are Graph-only** (not in the compose API) — setting them
is a Plane-B estate write that needs estate auth, not a `/`-on-the-draft. Recipients are writeable only
in compose/appointment drafts; on received items they are read-only references, not Graph-edit targets.
`/add-categories` on a _received_ item commits **immediately** (no on-send safety net) and the category
must already exist in the master list. The URI for `/add-attachment` is untrusted — it must be
allowlisted. Server-side `sendMail` and client auto-send are intentionally **not** modeled; the
reviewable path is `mail`/`compose`/`/compose-appointment` + the gate and the user's Send.
