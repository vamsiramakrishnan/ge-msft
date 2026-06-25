"""surface_cli.budget — the effect/read/cell budget check over an analyze() result."""

from .types import _cell_count


def _budget_exceeded(result, limits) -> bool:
    return (
        len(result["effects"]) > limits["max_effects"]
        or len(result["reads"]) > limits["max_reads"]
        or sum(_cell_count(e["range"]) for e in result["effects"]) > limits["max_cells"]
    )
