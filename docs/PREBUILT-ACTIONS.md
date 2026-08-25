# Prebuilt actions and query style

The task pane promotes three app-specific starters at all times. Each starter follows the same
shape: **outcome + object + scope + destination**. Labels stay short; prompts carry the grounding,
format, safety, and write-back contract.

| App | Promoted starters | Query style |
| --- | --- | --- |
| Word | Tighten / rewrite selection; Comment on issues; Review against… | Act on the selection or document, preserve meaning, and make proposed edits reviewable. |
| Excel | Create a chart; Summarize this range; Find anomalies / outliers | Name the selected range, the analytical outcome, and whether the result belongs in chat or the workbook. |
| PowerPoint | Draft a slide; Draft a section from the unit; Suggest a redesign | Name the slide or section, audience, takeaway, and expected destination. |
| OneNote | Summarize sources onto the page; Discover related sources; Write an audio overview script | Name the page or research unit and keep source provenance explicit. |
| Outlook | Brief my latest 10 emails; Prepare for my next meeting; Draft a reply | Name the mailbox or calendar window, requested fields, ordering, and evidence boundary. |
| Teams | Live notes & recap; Action items; Catch me up | Name the meeting or channel window and the expected recap, owners, and decisions. |

## Recommended prompt pattern

> Using **[grounded source]**, **[verb + outcome]** for **[scope/time window]**. Return **[specific
> fields or format]** in **[chat/host destination]**. If the required context is unavailable, state
> what is missing and do not infer or invent records.

Examples:

- Word: “Using this document, review the whole agreement against the master agreement. Add comments
  for material conflicts and cite the clause that triggered each finding.”
- Excel: “Using the selected range, find the five largest outliers. Return the row, metric, expected
  range, and likely explanation in chat; do not change cells.”
- PowerPoint: “Using the research unit, draft one executive slide on Q4 outlook. Lead with the
  takeaway, use at most three supporting points, and add the source for each claim.”
- OneNote: “Using this page and its linked sources, write a cited synthesis under a new summary
  heading. Separate confirmed facts from open questions.”
- Outlook: “Using connected Outlook mail context, brief the 10 most recent messages in received-date
  order. Return sender, subject, time, one-line summary, and required action. If mailbox context is
  unavailable, say so and do not invent messages.”
- Teams: “Using this meeting transcript, return decisions, action items with owner and due date, and
  unresolved questions. Mark missing owners or dates explicitly.”

## In-context commands

Word, Excel, and PowerPoint manifest context menus expose deterministic **Summarize** and **Explain**
commands for the live selection or range. They pass only a typed mode and re-ground the current
selection as `@this`; selected content is never persisted in the handoff. Outlook, OneNote, and
Teams use the always-visible task-pane starters because their Office host surfaces do not expose the
same selection-menu entry points through this unified manifest contract.
