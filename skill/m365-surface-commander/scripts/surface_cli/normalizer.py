"""surface_cli.normalizer — reorder a program into the OBSERVE→DERIVE→EFFECT→CONTROL normal form
(ADR-0008 §3), flagging a read-after-effect as a fresh-observation (VERIFY) boundary."""

import json

from .parser import _is_expr_line, parse_line, extract_command_block_meta, verified_frame_error

_READ_PHASE_VERBS = {
    "outline", "read", "search", "list", "inspect", "properties", "comments", "attachments",
    "tables", "slides", "neighbors", "context", "open",
    "workspace", "save", "cat", "grep",
}


def _phase_of(line: str) -> str:
    """Classify a program line into its canonical phase: OBSERVE / DERIVE / EFFECT / CONTROL."""
    rec = parse_line(line)
    if rec and rec.get("kind") == "analysis-binding":
        return "DERIVE"
    if rec and rec.get("kind") == "verified-finish":
        return "CONTROL"
    if _is_expr_line(line):
        return "DERIVE"  # a `let $x = …` binding or a bare pure pipeline
    if rec is None or "error" in rec:
        return "EFFECT"  # keep unknowns where the model put them (in the effect tail)
    verb = rec["verb"]
    if verb == "analyze":
        kind = json.loads(rec["request"]).get("kind")
        return "OBSERVE" if isinstance(kind, str) and kind in {"capture", "query", "reconcile", "inspect", "filter", "remove"} else "EFFECT"
    if verb in _READ_PHASE_VERBS:
        return "OBSERVE"
    if verb in ("done", "help"):
        return "CONTROL"
    return "EFFECT"  # every write verb + /<kind> invoke


def normalize(program_text: str):
    """Reorder a program into the OBSERVE -> DERIVE -> EFFECT -> CONTROL normal form (ADR-0008 §3),
    preserving the original order WITHIN each phase (binding and effect dependencies are
    order-sensitive). Returns (lines, notes). A read that appears AFTER an effect is a fresh-observation
    signal — it is kept in OBSERVE but a note flags that it may belong in a separate VERIFY turn."""
    inner, _closed = extract_command_block_meta(program_text)
    frame_error = verified_frame_error(program_text) if inner is not None else None
    if frame_error:
        # Do not turn an ambiguous response into a valid program by discarding other frames.
        return program_text.splitlines(), [frame_error]
    if inner is None:
        inner = program_text
    original = [line.strip() for line in inner.splitlines()
                if line.strip() and not line.strip().startswith("#")]
    parsed = [parse_line(line) for line in original]
    if any(rec and rec.get("kind") in ("analysis-binding", "verified-finish") for rec in parsed):
        # Phase sorting would hoist an inspect ahead of the artifact it references, or repair a
        # forbidden command after finish into an executable program. Preserve the program order.
        return original, ["Typed artifact programs retain dependency and completion order; run check before execution."]
    buckets = {"OBSERVE": [], "DERIVE": [], "EFFECT": [], "CONTROL": []}
    notes = []
    seen_effect = False
    for raw in inner.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        phase = _phase_of(line)
        if phase == "EFFECT":
            seen_effect = True
        elif phase == "OBSERVE" and seen_effect:
            notes.append(
                f"'{line}' reads after an effect — a read of post-write state belongs in a separate "
                "VERIFY turn (fresh observation), not this program."
            )
        buckets[phase].append(line)
    lines = buckets["OBSERVE"] + buckets["DERIVE"] + buckets["EFFECT"] + buckets["CONTROL"]
    return lines, notes
