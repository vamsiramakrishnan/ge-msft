---
title: Interactive Copy-Paste
kind: reference
skill: m365-release-operator
topics: [copy-paste, curl, har, widget-token, auth, secrets]
load_when: The task needs a browser-authenticated widget request, pasted cURL/HAR, device-code login, or another interactive user handoff.
---

# Interactive copy-paste

Use copy-paste only when the local API path cannot mint the required short-lived credential. The
goal is to let the user authenticate in the browser they already trust, then let repo tooling extract
only the minimum local token/config needed for the next command.

## Accepted paste inputs

- Chrome DevTools "Copy as cURL" for a `content-discoveryengine.googleapis.com` widget request.
- A HAR export containing the same request.
- A one-line cURL command pasted directly at a prompt.
- A multi-line cURL/HAR block terminated by `__GE_WIDGET_CURL_END__`.

Good widget inputs contain:

- `authorization: Bearer <widget JWT>`
- `x-server-token: <server token>`
- a JSON body with `configId`
- often one of `widgetListAvailableAgentViews`, `widgetStreamAssist`, `widgetCreateAgent`, or
  `files:upload`

## Harness pattern

```harness
intent   skills
mode     interactive
step     collect a fresh Gemini Enterprise widget request from the signed-in browser
paste    widget-curl sentinel=__GE_WIDGET_CURL_END__ save=/tmp/ge-widget-request.curl
run      scripts/update-ge-widget-skills.sh --credentials-file /tmp/ge-widget-request.curl --list-only
done
```

For replacement:

```harness
intent   skills
mode     interactive
step     replace the visible planner and surface commander widget skills with current local zips
paste    widget-curl sentinel=__GE_WIDGET_CURL_END__ save=/tmp/ge-widget-request.curl
run      scripts/update-ge-widget-skills.sh --credentials-file /tmp/ge-widget-request.curl
done
```

## Secret handling

- Store raw pasted requests in `/tmp` by default.
- Store extracted widget token in `/tmp/ge-widget-token` with mode `0600`.
- Store non-secret exports plus token-file pointer in `/tmp/ge-widget.env` with mode `0600`.
- Never persist bearer tokens, cookies, auth codes, or HAR files in git.
- Never paste the raw request back into the conversation after extraction.
- Redact tokens when summarizing failures.

## Token freshness

The harness should check whether `GE_WIDGET_BEARER_TOKEN_FILE` or `/tmp/ge-widget-token` exists and
has enough remaining lifetime. If it is fresh enough, skip browser paste. If it is missing, invalid,
or near expiry, ask for a new cURL/HAR paste.

This does not extend the widget token. It only avoids unnecessary re-auth while the token is still
valid.

## User experience

Keep the handoff short:

1. Tell the user exactly which browser request to copy.
2. Provide a single paste prompt with the sentinel.
3. Run the extractor locally.
4. Show only non-secret results: config id, visible skill labels, numeric agent ids, bundle hashes,
   and next command.
