"""Which names Power Query actually has.

Measured rather than guessed: ``scripts/build_m_inventory.py`` asks a
live engine for ``#shared``, the record of everything in scope in an M
expression, and commits the names.  That is what lets ``Table.Pivot``,
which exists, be reported as a gap in pyOpenVBA while
``Table.NoSuchFunction``, which does not, is M's own "the name wasn't
recognized" -- the same answer Power Query gives.
"""

from __future__ import annotations

from functools import lru_cache

from pyopenvba.mlang._inventory_data import NAMES as _RAW


@lru_cache(maxsize=1)
def known() -> frozenset[str]:
    """Every name the engine reported, lowercased for lookup."""
    return frozenset(one.lower() for one in _RAW.split(",") if one)


def has_name(name: str) -> bool:
    """Whether Power Query has a name of its own spelled this way."""
    return name.lower() in known()
