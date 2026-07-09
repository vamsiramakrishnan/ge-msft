"""surface_cli.generated_language — the verb/transform/kind tables, sourced from the GENERATED
language manifest (ADR-0008 §4) via the shared `parse_commands` loader. Not hand-maintained: the
manifest is emitted from @ge/contracts, so these tables can never diverge from the runtime grammar."""

import sys
from pathlib import Path

# The shared line parser (parse_commands.py) lives in the scripts/ dir — the package's parent.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from parse_commands import (  # noqa: E402
    TRANSFORM_NAMES,
    EFFECT_VERBS,
    ACTUATION_KINDS,
    COMMAND_HELP,
    CAPABILITY_REGISTRY,
    LANGUAGE_VERSION,
)

__all__ = [
    "TRANSFORM_NAMES",
    "EFFECT_VERBS",
    "ACTUATION_KINDS",
    "COMMAND_HELP",
    "CAPABILITY_REGISTRY",
    "LANGUAGE_VERSION",
]
