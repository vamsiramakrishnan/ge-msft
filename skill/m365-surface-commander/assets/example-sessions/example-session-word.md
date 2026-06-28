---
title: Example Session - Word
kind: example
skill: m365-surface-commander
surface: word
topics: [tracked-change, comment, claim-review]
load_when: A concrete Word redline/comment command-loop transcript is needed.
---

# Example session — Word

A review pass: find unsupported claims, propose tracked changes anchored on exact text, and
leave a comment. Shows content anchoring (not positions) and one write per line.

**User task:** "Tighten the opening and flag anything we can't source."

---

**Turn 1 — read the opening.**

````
**thought** Read the document so I can anchor edits on exact wording.
```cmd
outline
search "revenue grew"
```
````

The `result` returns the outline and the sentence containing the phrase, e.g.
_"Our revenue grew massively last year, the best ever."_

---

**Turn 2 — propose a tracked change and flag the unsourced claim.**

````
**thought** Replace vague wording with precise, and comment the unsourced superlative.
```cmd
suggest "grew massively last year, the best ever" => "grew 12% year over year"
comment "the best ever" "Unsupported superlative — cite a source or remove."
```
````

Each `suggest` becomes a tracked change anchored on the exact existing text (re-found at apply
time, so it survives minor edits). The comment is anchored on the quoted text. Both are
previewed, approved once, applied, and recorded.

---

**Turn 3 — finish.**

````
```cmd
done
```
````

---

### Notes

- `suggest` anchors on the **exact** text you read — never on a guessed position. If the text
  has changed, you get a correction and can re-read and retry.
- A claim you cannot support gets a `comment`, not a silent rewrite.
