"""Number format codes: as Range.NumberFormat spells them, and as the file does.

Setting Range.NumberFormat does not keep the code as written: Excel
parses it and writes it back out. What scripts/measure_typing_formats.py
and scripts/measure_format_codes.py measured of that:

- a quoted ``"$"`` becomes a bare ``$``, while ``"-"``, ``" "`` and other
  quoted text stay quoted; an escaped ``\\-``, ``\\(`` or ``\\)`` becomes
  bare, while ``\\$``, ``\\%`` and ``\\@`` stay escaped;
- date and time codes are lower case, ``General`` and colour names are
  capitalised, and ``am/pm`` becomes ``AM/PM`` while ``a/p`` stays;
- a lower-case exponent, ``0.00e+00``, is refused with error 1004.

The file spells a code another way. The bare ``$`` NumberFormat shows is
the currency symbol, which the file quotes as ``"$"``; a literal ``-``,
``(``, ``)`` or space is escaped with a backslash; a bare ``$`` in a file
is a plain dollar sign, which NumberFormat shows as ``\\$``. Excel's
built-in formats have ids of their own and are not written out, except
the currency and accounting ones, 5 to 8 and 41 to 44, whose code Excel
spells out in the file as well.

Excel also moves a thousands separator when a quoted literal splits the
digits -- ``#,##0"."00`` comes back ``###,0"."00`` -- which is not
modelled; such a code is kept as written.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from typing import Final

from pyopenvba.formula._display import tokens
from pyopenvba.interpreter._values import error

_COLOURS = {name.lower(): name for name in ("Black", "Blue", "Cyan", "Green", "Magenta", "Red", "White", "Yellow")}
_ELAPSED = re.compile(r"[hms]+", re.IGNORECASE)
#: Literals NumberFormat shows bare and the file escapes.
_ESCAPED_IN_FILE = "-() "

#: Excel's built-in formats on a US English system, by id, as NumberFormat spells them.
BUILTIN: Final[dict[int, str]] = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    5: "$#,##0_);($#,##0)", 6: "$#,##0_);[Red]($#,##0)", 7: "$#,##0.00_);($#,##0.00)",
    8: "$#,##0.00_);[Red]($#,##0.00)", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM",
    20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm",
    37: "#,##0_);(#,##0)", 38: "#,##0_);[Red](#,##0)", 39: "#,##0.00_);(#,##0.00)", 40: "#,##0.00_);[Red](#,##0.00)",
    41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', 42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', 44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
}
#: The built-in formats Excel still spells out in the file's numFmts.
SPELLED_OUT: Final = frozenset({5, 6, 7, 8, 41, 42, 43, 44})


def _rewritten(code: str, change: Callable[[str, str], str]) -> str:
    return "".join(change(kind, text) for kind, text in tokens(code))


def normalized(code: str, *, strict: bool = True) -> str:
    """A number format as Excel keeps it once Range.NumberFormat is set to ``code``.

    ``strict`` refuses what Excel refuses to set; a code read from a file
    is taken as it comes.
    """
    pieces: list[str] = []
    run: list[str] = []

    def close_run() -> None:
        # General and AM/PM take the case Excel writes them in; a/p keeps its own.
        text = re.sub(r"general", "General", "".join(run), flags=re.IGNORECASE)
        pieces.append(re.sub(r"am/pm", "AM/PM", text, flags=re.IGNORECASE))
        run.clear()

    for kind, text in tokens(code):
        if kind == "char":
            run.append(text.lower() if text in "YMDHS" else text)
            continue
        close_run()
        if kind == "quoted":
            pieces.append("$" if text == '"$"' else text)
        elif kind == "escaped":
            pieces.append(text[1] if text[1] in _ESCAPED_IN_FILE else text)
        elif kind == "bracket":
            inner = text[1:-1]
            if inner.lower() in _COLOURS:
                inner = _COLOURS[inner.lower()]
            elif _ELAPSED.fullmatch(inner):
                inner = inner.lower()
            pieces.append(f"[{inner}]")
        elif kind == "exponent" and text[0] == "e" and strict:
            raise error(1004, "Unable to set the NumberFormat property of the Range class")
        else:
            pieces.append(text)
    close_run()
    return "".join(pieces)


def from_file(code: str) -> str:
    """A code as a file spells it, as Range.NumberFormat reads it."""
    return normalized(_rewritten(code, lambda kind, text: "\\$" if kind == "char" and text == "$" else text),
                      strict=False)


def to_file(code: str) -> str:
    """A code as Range.NumberFormat spells it, as the file spells it."""

    def spelled(kind: str, text: str) -> str:
        if kind != "char":
            return text
        if text == "$":
            return '"$"'
        return "\\" + text if text in _ESCAPED_IN_FILE else text

    return _rewritten(code, spelled)
