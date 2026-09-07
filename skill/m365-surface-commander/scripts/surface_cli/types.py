"""surface_cli.types — the value-type constants and the A1 range algebra used for dependency and
cell inference. Pure; no I/O, no parser dependency."""

import re

# Transform → output value type (ADR-0008 §1 algebra). Aggregations collapse a Table to a Number;
# everything else stays a Table. A source (read/search/outline/$var) is a Table.
_AGG = {"sum", "avg", "min", "max", "count"}

_RANGE = re.compile(r"^(?:(?P<sheet>[^!]+)!)?(?P<a>[A-Za-z]+\d+)(?::(?P<b>[A-Za-z]+\d+))?$")
_CELL = re.compile(r"^([A-Za-z]+)(\d+)$")


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _parse_range(ref: str):
    """`Sheet!A1:G11` / `A1:G11` / `A1` → (sheet, c1, r1, c2, r2) in 1-based ints, or None."""
    m = _RANGE.match(ref.strip())
    if not m:
        return None
    a = _CELL.match(m.group("a"))
    if not a:
        return None
    c1, r1 = _col_to_num(a.group(1)), int(a.group(2))
    if m.group("b"):
        b = _CELL.match(m.group("b"))
        if not b:
            return None
        c2, r2 = _col_to_num(b.group(1)), int(b.group(2))
    else:
        c2, r2 = c1, r1
    sheet = (m.group("sheet") or "").strip().strip("'")
    return (sheet, min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))


def _overlap(x, y) -> bool:
    """Two parsed ranges intersect (same sheet + box overlap). A blank sheet matches any sheet."""
    if x is None or y is None:
        return False
    sx, x1, y1, x2, y2 = x
    sy, a1, b1, a2, b2 = y
    if sx and sy and sx.lower() != sy.lower():
        return False
    return not (x2 < a1 or a2 < x1 or y2 < b1 or b2 < y1)


def _cell_count(rng) -> int:
    if rng is None:
        return 0
    _s, c1, r1, c2, r2 = rng
    return (c2 - c1 + 1) * (r2 - r1 + 1)
