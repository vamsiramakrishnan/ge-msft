# Gemini Enterprise Microsoft 365 task pane

## Product Register

Product UI. This is an authenticated Microsoft 365 add-in task pane and command surface, not a marketing site.

## Users

Knowledge workers and tenant administrators using Word, Excel, PowerPoint, Outlook, OneNote, and Teams inside a controlled enterprise tenant. They are already in a document, workbook, deck, message, note, or meeting context and need grounded assistance without losing trust in the host surface.

## Core Purpose

Help the user ask grounded questions and stage reviewable document actions against the active Microsoft 365 host. The add-in must make capability, grounding, approval, provenance, and failure state visible enough to trust while keeping routine controls quiet until needed.

## Product Personality

Calm, precise, enterprise-grade. The interface should feel like a serious Office-native command surface: dense where useful, quiet by default, clear under scrutiny, and never decorative for its own sake.

## Trust Boundaries

Host content, connector content, model output, and user-authored prompt text are untrusted. The UI must reinforce that writes are gated, previewed, and explicitly approved. Diagnostics and failure messages should be actionable without exposing document content, workbook values, mail content, tokens, or tenant secrets.

## Interaction Principles

- Keep the active host and current grounding scope as the first visible facts.
- Show primary contextual actions immediately; move secondary controls into hover, focus, or explicit edit states.
- Use progressive disclosure for catalog routing, session skills, connector detail, citations, provenance, and command traces.
- Prefer stable rows, compact controls, and standard affordances over large cards.
- Make disabled, busy, approval, and unsupported states unmistakable without loud decoration.
- Preserve keyboard access for every hover-disclosed control.

## Anti-References

Avoid glossy AI chat demos, gradient-heavy SaaS shells, nested cards, oversized hero-like panels, decorative glass effects, novelty icons, and copy that explains the UI instead of helping the user act.

## Accessibility Baseline

WCAG AA contrast, visible focus, keyboard access for every interactive control, reduced-motion compliance, and no hover-only functionality without a focus path.
