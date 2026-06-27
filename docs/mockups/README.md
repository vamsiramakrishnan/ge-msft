# Mockups — the clickable UX spec

Self-contained, animated HTML prototypes of the add-in. Each is the Gemini Enterprise task pane
docked beside the host app, built in the shared visual language (the same gradient, chips, and
provenance affordances the real `web-shell` renders). Open any file in a browser; they simulate the
key interaction with no network.

| File | Surface | Shows |
|---|---|---|
| `0-word-deep.html` | Word (deep) | the full review experience — inline finding cards, guided walkthrough, citations |
| `1-word.html` | Word | contract review: inline tracked-change suggestion + the research-unit chips |
| `2-excel.html` | Excel | `=GE.ASK()` streaming into a cell + linked entity cards |
| `3-powerpoint.html` | PowerPoint | deck composer: draft a section, generate speaker notes, suggest a redesign |
| `4-onenote.html` | OneNote | synthesize curated sources onto the page with a tag per claim; audio overview |
| `5-teams.html` | Teams | in-meeting agent: live notes, action items, a reviewable recap |
| `6-command-pane.html` | (surface-agnostic) | the **`/` verbs + `@` mentions** command pane and what one line compiles to |

## Screenshots

`screenshots/` holds 2× PNG renders of each mockup (captured headless), so the UX is reviewable
without opening a browser:

```
screenshots/word.png  excel.png  powerpoint.png  onenote.png  teams.png  word-deep.png
screenshots/command-pane.png
```

## How the pane maps to the build

The mockups are the *intended interaction*, not the wired client — read them before building a
surface. The command pane (`6-command-pane.html`) is the human layer over the capability stack:

- **`/`** → an Intent verb (ADR-0005 skills / the `CapabilityManifest`), routed to the assist loop
  or an A2A specialist.
- **`@`** → a grounding source mapped to real `streamAssist` fields (`query.parts[]`,
  `toolsSpec.dataStoreSpecs`, `fileIds`).
- the grammar is carried into Gemini Enterprise by the two skills in `../../skill/`
  (`m365-command-planner` → ` ```plan `, `m365-surface-commander` → ` ```cmd `), mounted via
  `skillsSpec` — see `../api/discoveryengine/skills-and-agents.md`.
