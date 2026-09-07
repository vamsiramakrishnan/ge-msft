"""surface_cli — the deterministic PREFLIGHT compiler for m365-cli programs (ADR-0008 §4), split
into focused modules: types (range algebra), generated_language (manifest-sourced tables), parser,
checker (analyze), normalizer, budget. The runnable CLI entry is ../surface_cli.py.

It is a PURE compiler tool: it NEVER calls Office.js or Graph, acquires tokens, discovers
capabilities, executes code, or mutates anything. It loads generated vocabulary and advisory metadata; structural parser behavior is
checked against the production contracts by executable parity fixtures."""

from .checker import analyze
from .normalizer import normalize, _phase_of
from .budget import _budget_exceeded
from .types import _cell_count, _parse_range, _overlap
from .generated_language import LANGUAGE_VERSION

__all__ = [
    "analyze",
    "normalize",
    "_phase_of",
    "_budget_exceeded",
    "_cell_count",
    "_parse_range",
    "_overlap",
    "LANGUAGE_VERSION",
]
