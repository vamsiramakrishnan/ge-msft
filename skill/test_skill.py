#!/usr/bin/env python3
"""
Iterative, multi-surface test harness for the m365-surface-commander skill.

Drives streamAssist with ONLY our skill connected (empty toolsSpec → no web grounding, no data
stores) and simulates the Office add-in's multi-turn loop against a realistic mock document
(Excel analysis / Outlook thread / Word contract). Uses the robust reader in de_stub so thoughts,
citations, and suggestions are handled the way a real add-in must.

Modes:
  live  (default): calls the real streamAssist on the env-configured engine (GE_PROJECT/GE_ENGINE).
                   Asserts isolation (no grounding).
  --stub CANNED  : drives the SAME loop against canned replies from de_stub (offline, deterministic)
                   to validate the harness/parser without the API.

Usage:
  python3 test_skill.py --agent m365-surface-commander --surface excel
  python3 test_skill.py --agent m365-surface-commander --surface contract --raw
  python3 test_skill.py --stub                       # offline harness self-check
"""
import argparse
import json
import os
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "m365-surface-commander" / "scripts"))
from parse_commands import parse_block          # noqa: E402
from de_stub import read_response, make_response  # noqa: E402
from fixtures import FIXTURES, TASKS             # noqa: E402

# Live mode targets a real tenant and has NO baked-in defaults: GE_PROJECT, GE_PROJECT_NUMBER and
# GE_ENGINE must be set explicitly (the offline --stub self-check needs none of them).
REQUIRED_ENV = ("GE_PROJECT", "GE_PROJECT_NUMBER", "GE_ENGINE")
LOCATION = os.environ.get("GE_LOCATION", "global")
HTTP_TIMEOUT = 120  # seconds


def resolve_live_config():
    """Return (project, base_resource, url) from env, or exit clearly if a required var is unset."""
    missing = [name for name in REQUIRED_ENV if not os.environ.get(name)]
    if missing:
        sys.exit(
            "Refusing live mode: missing required environment variable(s): "
            f"{', '.join(missing)}. Set GE_PROJECT, GE_PROJECT_NUMBER and GE_ENGINE "
            "(or use --stub for the offline self-check)."
        )
    project = os.environ["GE_PROJECT"]
    project_number = os.environ["GE_PROJECT_NUMBER"]
    engine = os.environ["GE_ENGINE"]
    base = (f"projects/{project_number}/locations/{LOCATION}/collections/default_collection"
            f"/engines/{engine}/assistants/default_assistant")
    url = f"https://discoveryengine.googleapis.com/v1alpha/{base}:streamAssist"
    return project, base, url


def token():
    import google.auth, google.auth.transport.requests
    c, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
    c.refresh(google.auth.transport.requests.Request())
    return c.token


def call_live(tok, project, base, url, agent, query, session=None):
    body = {
        "query": {"parts": [{"text": query}]},
        "skillsSpec": {"skills": [{"name": f"{base}/agents/{agent}"}]},
        "toolsSpec": {},  # ISOLATION
    }
    if session:
        body["session"] = session
    req = urllib.request.Request(url, data=json.dumps(body).encode(),
                                 headers={"Authorization": "Bearer " + tok,
                                          "X-Goog-User-Project": project,
                                          "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
        return json.load(resp)


# Exact per-verb usage the add-in can inject each turn (highest-leverage anti-drift lever).
VERB_USAGE = {
    "outline": "outline",
    "read": "read <selector>          e.g. read Sales!A2:D9",
    "search": "search <text>",
    "set": 'set <A1> <value|=formula>  e.g. set F2 =SUMIF(A2:A9,"East",C2:C9)',
    "format": "format <range> k=v ...  e.g. format A1:C1 bold=true",
    "comment": 'comment <selector> "text"  e.g. comment C8 "anomalous spike"',
    "suggest": 'suggest "old" => "new"  e.g. suggest "ninety (90) days" => "thirty (30) days"',
    "reply": 'reply <commentId> "text"',
    "mail": 'mail "body text"         e.g. mail "Thursday 3pm works — please send the deck."',
    "compose": 'compose "Subject" "body"',
    "slide": 'slide "Title" "bullet" ...',
    "page": 'page "Title" "body"',
    "post": 'post "text"',
    "done": "done",
}


def render_caps(fixture):
    verbs = [v.strip() for v in fixture.caps.split(",")]
    usage = dict(VERB_USAGE)
    if fixture.surface in ("word", "powerpoint", "onenote"):
        # content surfaces anchor a comment on EXACT existing text, not a cell address
        usage["comment"] = ('comment "exact existing text to anchor on" "your note"  '
                            'e.g. comment "automatically renews" "Risk: no cancellation window"')
    lines = [f"  {usage.get(v, v)}" for v in verbs]
    return (f'<capabilities surface="{fixture.surface}"> Use ONLY these verbs, exactly this syntax:\n'
            + "\n".join(lines) + "\n</capabilities>")


def run(get_chunks, fixture, task, max_turns=8, raw=False):
    """Shared loop. get_chunks(query, session) -> raw chunk list."""
    first = (f"{render_caps(fixture)}\n{fixture.doc_state()}\nTask: {task}")
    session, query = None, first
    m = {"turns": 0, "cmd_blocks": 0, "prose_only": 0, "errors": 0,
         "grounding_leak": False, "citations": 0, "done": False}

    for turn in range(1, max_turns + 1):
        chunks = get_chunks(query, session)
        r = read_response(chunks)
        session = r["session"] or session
        m["turns"] = turn
        if r["citations"]:
            m["citations"] += len(r["citations"])
            m["grounding_leak"] = True  # in isolation mode any citation = a leaked data source
        parsed = parse_block(r["text"])
        cmds, has_block = parsed["commands"], parsed["block"] is not None
        tag = "[CMD]" if has_block else "[PROSE]"
        print(f"── turn {turn} {tag}"
              f"{'  ⚠GROUNDING' if r['citations'] else ''}"
              f"{('  thoughts='+str(len(r['thoughts']))) if r['thoughts'] else ''}")
        if raw or not has_block:
            print("   " + r["text"].strip()[:600].replace("\n", "\n   "))
        if not has_block:
            m["prose_only"] += 1
            query = ("No command block. Reply with EXACTLY one ```cmd block (flat command lines), "
                     "closed with ```. Do not answer in prose.")
            continue
        m["cmd_blocks"] += 1
        for c in cmds:
            if "error" in c:
                m["errors"] += 1
                print(f"   ✗ parse: {c['error']}")
            else:
                print("   • " + " ".join(f"{k}={v}" for k, v in c.items()))
        # apply every non-done command (even when `done` is batched in the same block) BEFORE exiting
        results = []
        for c in cmds:
            if c.get("verb") == "done":
                continue
            results.append(f"error: {c['error']}" if "error" in c else fixture.apply(c))
        if any(c.get("verb") == "done" for c in cmds):
            m["done"] = True
            print("── done")
            break
        query = "```result\n" + "\n".join(results) + "\n```"

    # surface-specific applied-effects summary
    print("\n=== applied effects ===")
    for attr in ("writes", "comments", "suggestions", "drafts"):
        vals = getattr(fixture, attr, None)
        if vals:
            print(f"   {attr}: {vals}")
    print("=== metrics ===")
    print(json.dumps(m, indent=2))
    return m


def stub_self_check():
    """Offline: feed a scripted, correct multi-turn solution through the loop to prove the harness
    (reader + parser + fixture apply) is sound — independent of the model."""
    fixture = FIXTURES["excel"]()
    scripted = iter([
        # turn 1: read + write all three totals (with thoughts/citations/suggestions complications)
        "```cmd\nread Sales!A2:D9\nset Sales!F2 =SUMIF(A2:A9,\"East\",C2:C9)\n"
        "set Sales!F3 =SUMIF(A2:A9,\"West\",C2:C9)\nset Sales!F4 =SUMIF(A2:A9,\"North\",C2:C9)\n"
        "comment Sales!C8 \"Anomalous spike — verify before reporting\"\n```",
        # turn 2: done
        "```cmd\ndone\n```",
    ])

    def get_chunks(query, session):
        text = next(scripted)
        return make_response(text, thoughts=["Reasoning about totals"],
                             citations=[{"title": "internal", "uri": "x", "content": "…"}],
                             suggestions=["Chart it?"], split_fence=True)

    print("### STUB self-check (offline) — surface=excel\n")
    return run(get_chunks, fixture, TASKS["excel"], max_turns=4, raw=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--agent")
    ap.add_argument("--surface", default="excel", choices=list(FIXTURES))
    ap.add_argument("--max-turns", type=int, default=8)
    ap.add_argument("--raw", action="store_true")
    ap.add_argument("--stub", action="store_true", help="offline harness self-check (no API)")
    args = ap.parse_args()

    if args.stub:
        stub_self_check()
        return
    if not args.agent:
        sys.exit("--agent required for live mode (or use --stub)")

    project, base, url = resolve_live_config()
    tok = token()
    fixture = FIXTURES[args.surface]()
    task = TASKS[args.surface]
    print(f"### LIVE  agent={args.agent}  surface={args.surface}\n### TASK: {task}\n")
    run(lambda q, s: call_live(tok, project, base, url, args.agent, q, s), fixture, task,
        max_turns=args.max_turns, raw=args.raw)


if __name__ == "__main__":
    main()
