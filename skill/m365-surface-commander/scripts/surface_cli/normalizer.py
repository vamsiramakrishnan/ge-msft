"""surface_cli.normalizer — reorder a program into the OBSERVE→DERIVE→EFFECT→CONTROL normal form
(ADR-0008 §3), flagging a read-after-effect as a fresh-observation (VERIFY) boundary."""

from .parser import _is_expr_line, parse_line, extract_command_block_meta

_READ_PHASE_VERBS = {
    "outline", "read", "search", "list", "inspect", "properties", "comments", "attachments",
    "tables", "slides", "neighbors", "context", "open",
}


def _phase_of(line: str) -> str:
    """Classify a program line into its canonical phase: OBSERVE / DERIVE / EFFECT / CONTROL."""
    if _is_expr_line(line):
        return "DERIVE"  # a `let $x = …` binding or a bare pure pipeline
    rec = parse_line(line)
    if rec is None or "error" in rec:
        return "EFFECT"  # keep unknowns where the model put them (in the effect tail)
    verb = rec["verb"]
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
    if inner is None:
        inner = program_text
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
