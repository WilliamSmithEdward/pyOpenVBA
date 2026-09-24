"""How many arguments Excel takes in a call of each of its functions.

Excel will not take a formula (error 1004) that calls a function with
too few arguments or too many. Each of its 525 functions was written
with 0 to 16 and 249 to 256 arguments in live Excel
(scripts/measure_formula_refusals.py,
tests/fixtures/formula_refusals.json), and the counts it took follow
from the formula engine's own, with these rules:

- a function newer than Excel 2007 takes at most 254, the file giving
  the 255th place to its name behind _xlfn.;
- a function taking its last arguments in pairs, SUMIFS's range and
  criterion, takes only whole pairs, SWITCH excepted, whose default
  comes alone at the end;
- a few take fewer or more than the engine's, and the functions the
  engine has not got take what :data:`_ELSEWHERE` says;
- a function Excel has not got, a macro's, or a name holding a LAMBDA
  takes up to 254, as any call does.
"""

from __future__ import annotations

from typing import Final

from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc.registry import FUNCTIONS
from pyopenvba.formula._prefixes import newer

#: The most arguments any call takes, and a call Excel writes with its name as the first of the file's 255.
_MOST: Final = 255
_MOST_NAMED: Final = 254
#: Where Excel takes other counts than the engine: the fewest and the most.
_OTHERWISE: Final[dict[str, tuple[int, int]]] = {
    "INDEX": (2, 4), "MAP": (2, 254), "BYCOL": (1, 2), "BYROW": (1, 2),
}
#: The functions the engine has not got: the fewest and the most arguments Excel takes.
_ELSEWHERE: Final[dict[str, tuple[int, int]]] = {
    "BAHTTEXT": (1, 1), "CUBEKPIMEMBER": (3, 4), "CUBEMEMBER": (2, 3), "CUBEMEMBERPROPERTY": (3, 3),
    "CUBERANKEDMEMBER": (3, 4), "CUBESET": (2, 5), "CUBESETCOUNT": (1, 1), "CUBEVALUE": (1, 255),
    "DETECTLANGUAGE": (1, 1), "FILTERXML": (2, 2), "FORECAST.ETS": (3, 6), "FORECAST.ETS.CONFINT": (3, 7),
    "FORECAST.ETS.SEASONALITY": (2, 4), "FORECAST.ETS.STAT": (3, 6), "GETPIVOTDATA": (2, 254), "GROUPBY": (3, 8),
    "IMAGE": (1, 5), "INFO": (1, 1), "ODDFPRICE": (8, 9), "ODDFYIELD": (8, 9), "ODDLPRICE": (7, 8),
    "ODDLYIELD": (7, 8), "PHONETIC": (1, 1), "PIVOTBY": (4, 11), "RTD": (3, 255), "STOCKHISTORY": (2, 254),
    "TRANSLATE": (1, 3), "WEBSERVICE": (1, 1),
}
#: Functions a formula a macro writes cannot call at all: CALL and REGISTER.ID are XLM's, PY Python's own cell.
_NEVER: Final = frozenset({"CALL", "PY", "REGISTER.ID"})


def takes(name: str, count: int) -> bool:
    """Whether Excel takes a call of ``name``, in capitals as a formula spells it, with ``count`` arguments."""
    if name in _NEVER:
        return False
    if name == "GETPIVOTDATA":
        # The data field and the pivot table, then fields and items in pairs; or one more, as Excel 2000 took it.
        return count in (2, 3) or 4 <= count <= 254 and count % 2 == 0
    if name in _ELSEWHERE:
        fewest, most = _ELSEWHERE[name]
        return fewest <= count <= most
    entry = FUNCTIONS.get(name)
    if entry is None:
        return count <= _MOST_NAMED
    fewest, most = _OTHERWISE.get(name, (entry.minimum, entry.maximum))
    most = min(most, _MOST_NAMED if newer(name) else _MOST)
    if not fewest <= count <= most:
        return False
    declared = len(entry.kinds)
    return entry.repeat == 1 or name == "SWITCH" or count < declared or (count - declared) % entry.repeat == 0


__all__ = ["takes"]
