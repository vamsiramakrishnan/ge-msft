#!/usr/bin/env python3
"""
Stub + reader for the Discovery Engine `streamAssist` response, matching the real wire shape we
observed on phoenix-telco:

  response = [ chunk, chunk, ... ]
  chunk = {
    "answer": { "replies": [ { "replyId", "createTime", "groundedContent": {
                  "content": { "role": "model", "text": "<token>", "thought"?: bool,
                               "inlineData"?: { "mimeType", "data(base64)" } },
                  "textGroundingMetadata"?: { "references": [ { "content", "title"?, "uri"? } ] }
               } } ] },
    "sessionInfo": { "session", "queryId" },
    "assistToken": "<opaque>"
  }

Real complications this reproduces, so the reader is exercised against them:
  - text is streamed token-by-token across MANY chunks (must be reassembled in order)
  - "thoughts": the thinking model emits bold-header prose (and may set content.thought=true)
  - citations live in textGroundingMetadata.references (not inline in the text)
  - inlineData carries application/json+suggestions (recommended follow-ups) — not answer text
  - a ```cmd block can be SPLIT across chunks (incl. the closing fence arriving late)

Two entry points:
  make_response(model_text, *, thoughts=None, citations=None, suggestions=None, session=...)
      -> a realistic chunk list (for offline tests)
  read_response(chunks)
      -> { "text", "thoughts", "citations", "suggestions", "session", "queryId" }
      the robust reader the add-in/harness should use (skips thoughts & inlineData from answer
      text, collects citations and suggestions separately).
"""
from __future__ import annotations

import base64
import json


def _chunk(reply_content, *, tgm=None, session="stub/sessions/1", qid="q1", rid="r"):
    gc = {"content": reply_content}
    if tgm:
        gc["textGroundingMetadata"] = tgm
    return {
        "answer": {"replies": [{"replyId": rid, "createTime": "2026-06-24T00:00:00Z",
                                "groundedContent": gc}]},
        "sessionInfo": {"session": session, "queryId": qid},
        "assistToken": "STUBTOKEN",
    }


def _tokenize(text, n=6):
    """Split text into ~n-char tokens to emulate streaming."""
    return [text[i:i + n] for i in range(0, len(text), n)] or [""]


def make_response(model_text, *, thoughts=None, citations=None, suggestions=None,
                  session="stub/sessions/1", split_fence=False):
    """Build a realistic chunk list. `model_text` is the FULL reply (may contain a ```cmd block)."""
    chunks = []
    rid = 0

    # 1) thoughts first (as the thinking model would surface them) — bold headers, thought=true
    for th in (thoughts or []):
        rid += 1
        chunks.append(_chunk({"role": "model", "text": f"**{th}**", "thought": True},
                             session=session, rid=f"t{rid}"))

    # 2) the answer text, streamed token-by-token (optionally split right before the closing fence)
    body = model_text
    pieces = []
    if split_fence and "```" in body:
        # break so the final ``` arrives as its own late chunk
        head, _, _ = body.rpartition("```")
        pieces = _tokenize(head) + ["```"]
    else:
        pieces = _tokenize(body)
    for i, tok in enumerate(pieces):
        rid += 1
        # attach citations to the last answer chunk (like the real API trails grounding metadata)
        tgm = None
        if citations and i == len(pieces) - 1:
            tgm = {"references": [{"content": c.get("content", ""),
                                   **({"title": c["title"]} if c.get("title") else {}),
                                   **({"uri": c["uri"]} if c.get("uri") else {})}
                                  for c in citations]}
        chunks.append(_chunk({"role": "model", "text": tok}, tgm=tgm, session=session, rid=f"a{rid}"))

    # 3) suggestions as inlineData (application/json+suggestions)
    if suggestions:
        payload = {"recommendedQuestionsResponse": {"questions": suggestions}}
        data = base64.b64encode(json.dumps(payload).encode()).decode()
        rid += 1
        chunks.append(_chunk({"role": "model",
                              "inlineData": {"mimeType": "application/json+suggestions", "data": data}},
                             session=session, rid=f"s{rid}"))
    return chunks


def read_response(chunks):
    """Robustly read a streamAssist chunk list. Separates answer text, thoughts, citations,
    suggestions — the way a correct add-in/harness must."""
    text_parts, thoughts, citations, suggestions = [], [], [], []
    session = queryId = None
    seen_ref = set()
    for c in chunks:
        si = c.get("sessionInfo") or {}
        session = si.get("session", session)
        queryId = si.get("queryId", queryId)
        for r in c.get("answer", {}).get("replies", []):
            gc = r.get("groundedContent", {})
            content = gc.get("content", {})
            # suggestions / non-text payloads
            idt = content.get("inlineData")
            if idt and idt.get("mimeType", "").endswith("+suggestions"):
                try:
                    dec = json.loads(base64.b64decode(idt["data"]).decode())
                    suggestions.extend(dec.get("recommendedQuestionsResponse", {}).get("questions", []))
                except Exception:
                    pass
                continue
            # thoughts (flagged, or bold-header prose) are NOT answer text
            txt = content.get("text")
            if txt is None:
                pass
            elif content.get("thought"):
                thoughts.append(txt.strip("* "))
            else:
                text_parts.append(txt)
            # citations
            tgm = gc.get("textGroundingMetadata") or {}
            for ref in tgm.get("references", []):
                key = (ref.get("title"), ref.get("uri"), ref.get("content", "")[:40])
                if key in seen_ref:
                    continue
                seen_ref.add(key)
                citations.append({"title": ref.get("title"), "uri": ref.get("uri"),
                                  "content": ref.get("content", "")})
    return {"text": "".join(text_parts), "thoughts": thoughts, "citations": citations,
            "suggestions": suggestions, "session": session, "queryId": queryId}


# ───────────────────────── self-test ─────────────────────────
if __name__ == "__main__":
    model = "**thinking**\nLet me total East.\n```cmd\nread Sales!A2:C9\nset Sales!F2 =SUMIF(A2:A9,\"East\",C2:C9)\n```"
    chunks = make_response(
        model,
        thoughts=["Calculating East Region Total", "Choosing SUMIF"],
        citations=[{"title": "RackNap blog", "uri": "https://example.com/telco", "content": "Telecom bundling is evolving…"}],
        suggestions=["How do telcos split revenue?", "Show B2B examples"],
        split_fence=True,
    )
    print(f"made {len(chunks)} chunks")
    out = read_response(chunks)
    print("text reassembled:\n", out["text"])
    print("thoughts:", out["thoughts"])
    print("citations:", [(c["title"], c["uri"]) for c in out["citations"]])
    print("suggestions:", out["suggestions"])
    # prove the (split) fence still extracts
    import sys, pathlib
    sys.path.insert(0, str(pathlib.Path(__file__).parent / "m365-surface-commander" / "scripts"))
    from parse_commands import parse_block
    print("parsed commands:", [c.get("verb", c) for c in parse_block(out["text"])["commands"]])
