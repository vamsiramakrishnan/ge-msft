# Gemini Enterprise — Microsoft 365 task pane

## Register

product

## Users

Knowledge workers and tenant administrators using Word, Excel, PowerPoint, Outlook,
OneNote, and Teams inside a controlled enterprise tenant. They are already in a document,
workbook, deck, message, note, or meeting context and need grounded assistance without
losing trust in the host surface. They arrive mid-task, not at a blank page, and expect the
pane to respect the work already on screen.

## Product Purpose

Help the user ask grounded questions and stage reviewable document actions against the active
Microsoft 365 host. Success is the user trusting the pane enough to act through it: capability,
grounding scope, approval, provenance, and failure state are all visible enough to trust, while
routine controls stay quiet until needed. The pane is a command surface embedded in the host,
not a standalone chatbot — every write is gated, previewed, and explicitly approved before it
touches the document.

## Brand Personality

Calm, precise, enterprise-grade. The interface should read like a serious Office-native command
surface: dense where useful, quiet by default, clear under scrutiny, and never decorative for its
own sake. Voice is plain and operational — it helps the user act rather than narrating the UI.
The Gemini identity is present as a small product mark; the host (Word/Excel/…) owns the
interaction color so the pane feels native to wherever it is docked.

## Anti-references

Avoid glossy AI-chat demos, gradient-heavy SaaS shells, nested cards, oversized hero-like
panels, decorative glass effects, novelty icons, tracked-uppercase eyebrows on every section,
and copy that explains the UI instead of helping the user act. No marketing-site rhetoric — this
is an authenticated tool, not a landing page.

## Design Principles

- **Host first, tool second.** The active host and current grounding scope are the first visible
  facts, and the host's accent is the interaction color. The pane should feel like it belongs to
  Word or Excel, not like a guest window. The tool disappears into the task.
- **Quiet by default, dense on demand.** Primary contextual actions are immediate; catalog
  routing, session skills, connector detail, citations, provenance, and command traces stay
  behind hover/focus/edit disclosure. Reach for stable rows and compact controls over large cards.
- **Every write is earned.** Prefer reversible, provenanced actions over silent edits. Writes are
  dry-run, previewed, and explicitly approved through a visible gate; the gate is a feature, not
  friction to hide.
- **Untrusted until proven.** Host content, connector content, model output, and user-typed prompt
  text are all untrusted. The chrome must continuously signal that writes are staged and approved,
  and diagnostics must be actionable without leaking document content, values, tokens, or secrets.
- **Legible under stress.** Disabled, busy, approval, unsupported, and degraded states must be
  unmistakable without loud decoration, and every hover affordance must also open through keyboard
  focus. Clarity when something is wrong matters more than polish when everything is fine.

## Trust Boundaries

Host content, connector content, model output, and user-authored prompt text are untrusted. The
UI must reinforce that writes are gated, previewed, and explicitly approved. Diagnostics and
failure messages should be actionable without exposing document content, workbook values, mail
content, tokens, or tenant secrets.

## Interaction Principles

- Keep the active host and current grounding scope as the first visible facts.
- Show primary contextual actions immediately; move secondary controls into hover, focus, or
  explicit edit states.
- Use progressive disclosure for catalog routing, session skills, connector detail, citations,
  provenance, and command traces.
- Prefer stable rows, compact controls, and standard affordances over large cards.
- Make disabled, busy, approval, and unsupported states unmistakable without loud decoration.
- Preserve keyboard access for every hover-disclosed control.

## Accessibility & Inclusion

WCAG AA contrast (body text ≥4.5:1, large/bold ≥3:1, placeholders held to the same body
standard), visible focus on every interactive control, full keyboard access with a focus path
for every hover-disclosed control (no hover-only functionality), and `prefers-reduced-motion`
compliance. The pane is designed from the narrow Office pane outward, so controls must stay
reachable as width and height compress.
