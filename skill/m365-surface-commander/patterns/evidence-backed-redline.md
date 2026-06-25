# Pattern: evidence-backed redline

A reasoning template, not a command. Read it for shape, then write the turn's algebra.

**Intent:** "redline this document — fix claims, flag the unsourced ones."

**Preconditions**

- the source supports `read`/`search`;
- the target supports `suggest` (tracked change) and `comment`.

**Pure core (OBSERVE → DERIVE)**

```
read                       # the whole document (or search for the claim-bearing passages)
search "<claim phrase>"    # locate the exact text to anchor on
  → reason over what you READ: which claims are wrong, which are unsourced
```

You cannot anchor on text you have not read — `suggest` and `comment` both require the **exact
existing text**. Read first, then anchor.

**Effect core (EFFECT)**

```
suggest "exact old wording" => "corrected wording"     # a reversible tracked change, anchored
comment "exact unsourced sentence" "needs a citation"  # flag, don't silently change
```

A correction the source supports → `suggest` (tracked change). A claim you cannot verify → `comment`
(flag for the human), never a silent rewrite.

**Failure rule:** if the exact anchor text has drifted (you can't find it), degrade to a panel
comment rather than emitting a broken annotation — never anchor on approximate text.

**Anti-patterns**

- rewriting an unsourced claim instead of flagging it;
- `suggest` with an anchor you guessed rather than read (it won't resolve);
- one giant tracked change over a whole paragraph when the fix is one phrase.
