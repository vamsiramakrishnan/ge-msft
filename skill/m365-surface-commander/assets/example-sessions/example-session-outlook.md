# Example session — Outlook

Draft a grounded reply to the open email. Shows reading the active item and staging a
reviewable reply — never auto-sent.

**User task:** "Reply to confirm the Thursday 3pm slot and ask them to send the agenda."

---

**Turn 1 — read the open mail item.**

````
**thought** Read the message so the reply matches what they actually asked.
```cmd
read
search "meeting"
```
````

The `result` returns the subject, sender, and leading body lines — the sender proposed a few
times including Thursday 3pm.

---

**Turn 2 — stage the reply.**

````
**thought** Confirm the slot, request the agenda, keep it short with a clear next step.
```cmd
mail "Thursday at 3pm works for me — I'll send a calendar invite. Could you share the agenda beforehand so I can prepare? Thanks!"
```
````

This opens a **reviewable reply draft** in Outlook. It is never sent automatically — the user
reviews and sends.

---

**Turn 3 — finish.**

````
```cmd
done
```
````

---

### Drafting a brand-new email

To start a fresh message instead of replying, use `compose`. Recipients are left blank for the
user to fill — never auto-addressed.

````
```cmd
compose "Q3 planning — agenda request" "Hi — ahead of Thursday's session, could you send the agenda and any pre-reads? I'll circulate notes afterward."
```
````
