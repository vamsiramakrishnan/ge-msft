# Flaky capabilities log

A **flaky** result is one that fails on one run and passes on a rerun with no code change.
We record every one here, so we can see patterns and do not mistake flakiness for a regression.

**How to add an entry**
- One row per scenario per day. If it happens again later, add a new row (the history matters).
- "Rerun" = the result of running only that scenario again (`GE_SCENARIO=<id> bun run scenarios`).
- "Cause" = what the transcript shows (model loop, timeout, wrong verb…). Write "unknown" if not checked.
- Raw results stay in `scenarios/.results/<run>/` (git-ignored; transcripts can hold test data).

**Known flaky unit test:** "a host snapshot that never settles is skipped after the read timeout"
(web-shell). Passes on rerun; use `--retry=1`.

## Patterns so far

| Pattern | Capabilities | First seen |
|---|---|---|
| Model loops on `read` / `search` and never writes | outlook:reply | 2026-10-15 |
| A read-only request opens a new mail draft | outlook:read | 2026-10-15 (also fixed once in `68c93c7`, 2026-10-13) |
| No answer for 90 s → timeout | outlook (attachment) | 2026-10-15 |

## Log

| Date | Capability | Scenario | Prompt | First run | Rerun | Cause (from transcript) | Change under test |
|---|---|---|---|---|---|---|---|
| 2026-10-15 | outlook:newmail | `mx-ol-newmail-04` | Draft an email to Budi asking for the payment gateway test results | blocked (29 s): program failed validation after 2 correction turns | pass | Model wrote an invalid command program | Outlook line breaks + remove "+ Sources" |
| 2026-10-15 | outlook:newmail | `mx-ol-newmail-08` | Write an email to the vendor asking for an update | completed, 0 × create-mail | pass | Model answered without opening a draft | same |
| 2026-10-15 | outlook:read | `mx-ol-read-09` | What are the risks mentioned? | blocked (25 s) | pass | unknown | same |
| 2026-10-15 | outlook:read | `mx-ol-read-10` | Extract the action items as a list | failed (97 s) | **fail**: create-mail applied 1×, expected none | Read-only request opened a new mail draft | same |
| 2026-10-15 | outlook:reply | `mx-ol-reply-10` | Draft a reply with the target dates for both checks | blocked (91 s) | pass | Model repeated `read` / `search` 8+ turns, never drafted | same |
| 2026-10-15 | outlook (drafts) | `ol-draft-attach` | Attach a text file named freeze-checklist.txt … | failed (122 s): no answer for 90 s | not rerun | Timeout | same |
| 2026-10-17 | excel:unsupported | `mx-xl-unsupported-02` | Buat pivot table outstanding per segment dan collectability | timed out: planner streamed only thoughts for ~10 min, no plan | completed but **approximated**: planner planned a "summarized pivot table", model wrote a typed summary grid | Gemini stall (run 1); planner plans an approximation (run 2) | honest-stop merge 9d48e0c/8d75496 |
