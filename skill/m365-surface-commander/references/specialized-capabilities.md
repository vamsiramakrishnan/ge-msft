# The `/` specialized-capability surface

There are two tiers of capability (ADR-0008 §two-tier):

1. **The composable algebra** — the small set of bare verbs (`read`/`filter`/`sort`/`head`/`sum` +
   `set`/`spill`/`table`/`chart`/`cf`/`suggest`/`comment`/`reply`). These compose into programs. Learn
   them well; they are the assembly language.
2. **The specialized catalogue** — the long tail of host-native capabilities (insert an image, attach a
   file, fill a content control, post to a channel, set a page title, …). These are **named, typed,
   non-composing effect terminals** reached with a leading `/`.

You don't memorize the catalogue. You reach a specialized capability by name with `/`.

## Syntax

```
/<kind> [positional …] [key=value …]
```

- **`<kind>` is the `ActuationKind` itself** — e.g. `/insert-image`, `/add-attachment`,
  `/fill-content-control`, `/post-channel-message`, `/set-page-title`. There is no separate alias to
  learn; the command name is the capability.
- Arguments use the same `key=value` (+ positional) grammar as `format`/`table`; quote values with
  spaces: `alt="Q3 revenue chart"`.
- A `/` command is an **effect terminal** — it does not compose, you cannot pipe out of it, and its
  result is not a value. (Use the algebra for composition; use `/` for a specialized leaf effect.)

## Discovery — only what's available this turn

A `/` capability is usable **only when this surface advertises that kind this turn.** The available set
is the catalogue (`scripts/m365-cli-1.0.json` → `specializedKinds`) intersected with this turn's
`<capabilities>` signature. Reaching for a `/<kind>` the surface doesn't handle is rejected
(`/<kind> is not supported on this surface`), exactly like an unavailable core verb. An unknown kind
name gets a catalogue did-you-mean (`unknown capability "/insert-imag" — did you mean "/insert-image"?`).

## Examples

```
# Word
/insert-image base64=<…> alt="Top regions by revenue"
/fill-content-control id=Title text="Q3 Business Review"
/insert-hyperlink url=https://… text="the source"

# OneNote
/set-page-title title="Meeting notes — 2026-06-25"
/add-note-tag type=toDo

# Outlook (on the open draft)
/add-attachment name="report.xlsx" base64=<…>

# Teams / estate (only when the estate-write capability is advertised)
/post-channel-message text="Summary posted to the deck channel"
```

## When to preflight

A `/` command is an effect — it counts toward the effect budget and is shown on the approval card. If
you emit more than a couple of effects, or a `/<kind>` that reaches the estate (a post, a calendar
event, a triage), run `surface_cli check` (it scopes the kind against the turn's capabilities and
flags external effects). See the algebra in [algebra.md](algebra.md) and the laws in
[composition-rules.md](composition-rules.md).
