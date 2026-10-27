"""surface_cli.budget — the effect/read/cell budget check over an analyze() result."""

from .types import _cell_count
from .generated_language import LIMITS

# The runtime's default per-turn limits, carried in the generated manifest (ADR-0008 "budget is
# defense-in-depth": the preflight advises against the same numbers the gate re-checks). The
# literal fallbacks only apply to a bundle that predates the `limits` block.
DEFAULT_LIMITS = {
    "max_effects": int(LIMITS.get("maxWritesPerTurn", 8)),
    "max_reads": int(LIMITS.get("maxReadsPerTurn", 8)),
    "max_cells": int(LIMITS.get("maxCellsPerTurn", 10000)),
}


def _budget_exceeded(result, limits) -> bool:
    return (
        len(result["effects"]) > limits["max_effects"]
        or len(result["reads"]) > limits["max_reads"]
        or sum(_cell_count(e["range"]) for e in result["effects"]) > limits["max_cells"]
    )
