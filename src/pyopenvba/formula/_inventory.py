"""Which worksheet functions Excel actually has.

The same question the object model asks about members, for the same
reason: a function Excel has and this does not is a gap in pyOpenVBA
and says so, while a name Excel has never had is ``#NAME?``, which is
what Excel itself shows.

Most of the list comes from the WorksheetFunction class in the type
library.  The rest are the ones a sheet has and VBA does not expose
there, which are listed here by name.
"""

from __future__ import annotations

from functools import lru_cache
from typing import Final

#: Functions a cell has that WorksheetFunction does not carry.
SHEET_ONLY: Final = frozenset(
    """
    CELL COLUMN COLUMNS INDIRECT INFO NOW OFFSET ROW ROWS TODAY RAND RANDBETWEEN RANDARRAY
    IF IFS IFERROR IFNA AND OR NOT TRUE FALSE SWITCH LET LAMBDA
    N T TYPE ERROR.TYPE NA ISBLANK ISERR ISERROR ISLOGICAL ISNA ISNONTEXT ISNUMBER ISREF ISTEXT
    ISFORMULA ISODD ISEVEN SHEET SHEETS FORMULATEXT
    SUM AVERAGE COUNT COUNTA MAX MIN TEXT VALUE LEN LEFT RIGHT MID TRIM UPPER LOWER
    CONCAT CONCATENATE TEXTJOIN EXACT FIND SEARCH SUBSTITUTE REPLACE REPT CHAR CODE UNICHAR UNICODE
    ABS SIGN INT TRUNC ROUND ROUNDUP ROUNDDOWN MOD POWER SQRT EXP LN LOG LOG10 PI
    DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND WEEKDAY DATEVALUE TIMEVALUE DAYS
    CHOOSE INDEX MATCH VLOOKUP HLOOKUP XLOOKUP XMATCH LOOKUP TRANSPOSE SORT SORTBY UNIQUE FILTER
    SEQUENCE TEXTSPLIT TEXTBEFORE TEXTAFTER VSTACK HSTACK TOCOL TOROW WRAPCOLS WRAPROWS
    TAKE DROP EXPAND CHOOSECOLS CHOOSEROWS ARRAYTOTEXT VALUETOTEXT
    """.split()
)


@lru_cache(maxsize=1)
def _known() -> frozenset[str]:
    from pyopenvba.interpreter._inventory import members_of

    # The members are on the interface; the class beside it is empty.
    found = members_of("WorksheetFunction", "excel") | members_of("IWorksheetFunction", "excel")
    return frozenset(name.upper() for name in found) | SHEET_ONLY


def excel_has_function(name: str) -> bool:
    """Whether Excel has a worksheet function of this name."""
    return name.upper() in _known()


def unimplemented() -> list[str]:
    """Every Excel function this does not implement, for the docs to list."""
    from pyopenvba.formula._functions import known_names

    return sorted(_known() - known_names())
