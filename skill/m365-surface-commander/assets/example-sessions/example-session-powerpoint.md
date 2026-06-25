# Example session — PowerPoint

Build a slide from content already in the deck: read the slide inventory, find the supporting
numbers, and insert a clean summary slide. Shows reading slides and composing bullets from a
value.

**User task:** "Add a summary slide with our top three regions by revenue."

---

**Turn 1 — see the deck, then read the data slide.**

````
**thought** Get the slide list, then read the slide that holds the regional numbers.
```cmd
outline
read slide:3
```
````

The `result` returns the slide inventory and the text of slide 3 (a table of regions and
revenue).

---

**Turn 2 — insert the summary slide.**

The top three from what was read are East, West, and North.

````
**thought** One idea per slide; short bullets; numbers I read, not guessed.
```cmd
slide "Top regions by revenue" "East — $4.2M" "West — $3.1M" "North — $2.7M"
```
````

The new slide is previewed, approved, inserted, and recorded.

---

**Turn 3 — finish.**

````
```cmd
done
```
````

---

### Composition variant

If the regional figures were on an Excel range earlier in the session, the bullets can be
composed from a value (rows become bullets):

````
```cmd
slide "Top regions by revenue" ($top3 | select region,revenue)
```
````

Keep each slide to a single message and a few short bullets.
