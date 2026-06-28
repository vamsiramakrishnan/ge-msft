---
title: Pattern - Meeting Summary To Notes
kind: pattern
skill: m365-surface-commander
surface: teams
topics: [meeting, transcript, notes, action-items]
load_when: A task asks to summarize a Teams meeting or turn transcript content into notes/actions.
---

# Pattern: meeting summary → notes

A reasoning template, not a command. Read it for shape, then write the turn's algebra.

**Intent:** "summarize this meeting and capture the action items."

**Preconditions**

- the source supports `read` (a transcript / the open notes);
- the target supports `page` (OneNote) or `post` (Teams).

**Pure core (OBSERVE → DERIVE)**

```
read                          # the transcript / meeting content
  → derive: a short summary + a list of decisions + a list of action items (owner, task)
```

Bind the derived pieces if you reuse them: `let $actions = read … | filter …` where the transcript is
tabular; otherwise reason over the read text directly.

**Effect core (EFFECT)**

```
# OneNote
page "Meeting notes — <date>" "<summary> … <decisions> … <action items>"

# Teams (reviewable, never auto-sent)
post "Summary: … · Actions: …"
```

One synthesized page/post, not a flurry of small writes. Action items that should become tasks are a
**separate, gated** step (a `/create-task` estate effect with its own approval) — don't fold an
external task-create into the same note write (approval-authority boundary).

**Failure rule:** the summary is derived from what you read; never attribute a decision or an owner the
transcript doesn't contain.

**Anti-patterns**

- inventing action items / owners not in the transcript;
- posting to Teams AND creating tasks under one approval (different failure domains);
- a dozen tiny `post`s instead of one structured summary.
