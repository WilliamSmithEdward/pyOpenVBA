"""Translate cell-relative formula notation without rewriting string literals."""
from __future__ import annotations

import re

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, column_letter, column_number
from pyopenvba.formula._parse import STRUCTURED, split_sheet, tokenize

_OFFSET = r"(?:\[-?\d+\]|\d+)?"
_ROW = rf"R{_OFFSET}"
_COL = rf"C{_OFFSET}"
_SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z0-9_.À-￿]+)!"
#: A structured reference comes after the R1C1 forms, so that R[-1]C reads as one; what is in its brackets, a
#: column named R1C1 or C, is left as it is.
_SCAN = re.compile(
    rf'(?P<text>"(?:[^"]|"")*")|(?P<ref>(?<![\w.])(?:{_SHEET})?'
    rf'(?:{_ROW}{_COL}(?::{_ROW}{_COL})?|{_ROW}:{_ROW}|{_COL}:{_COL}|{_ROW}|{_COL})(?![\w.(]))'
    rf"|(?P<sheet>{_SHEET})|(?P<structured>{STRUCTURED})", re.IGNORECASE,
)
_RC = re.compile(rf"^(?:R(?P<row>{_OFFSET}))?(?:C(?P<col>{_OFFSET}))?$", re.IGNORECASE)
_A1 = re.compile(r"^(\$?)([A-Za-z]+)?(\$?)(\d+)?$")


def _prefix(text: str, reference: str) -> str:
    prefix = text[:-len(reference)]
    if prefix.startswith("'"):
        sheet = prefix[1:-2].replace("''", "'")
        if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", sheet):
            return sheet + "!"
    return prefix


def to_a1(formula: str, row: int, column: int) -> str:
    """Resolve relative coordinates at a cell, wrapping at worksheet edges."""
    def replace(match: re.Match[str]) -> str:
        text = match.group()
        if match.lastgroup != "ref":
            return text
        _, reference = split_sheet(text)
        parts: list[str] = []
        for corner in reference.split(":"):
            rc = _RC.fullmatch(corner)
            assert rc is not None
            axes: list[str] = []
            for axis, origin, limit in (("col", column, MAX_COLUMNS), ("row", row, MAX_ROWS)):
                value = rc.group(axis)
                if value is None:
                    continue
                absolute = bool(value) and not value.startswith("[")
                number = int(value) if absolute else (origin + (int(value[1:-1]) if value else 0) - 1) % limit + 1
                if not 1 <= number <= limit:
                    raise ValueError("R1C1 reference outside worksheet")
                axes.append(("$" if absolute else "") + (column_letter(number) if axis == "col" else str(number)))
            parts.append("".join(axes))
        # A single R or C denotes the entire row or column.
        if len(parts) == 1 and ("R" not in reference.upper() or "C" not in reference.upper()):
            parts *= 2
        return _prefix(text, reference) + ":".join(parts)
    return _SCAN.sub(replace, formula)


def from_a1(formula: str, row: int, column: int) -> str:
    """Express A1 references relative to the cell containing the formula."""
    pieces: list[tuple[int, int, str]] = []
    for token in tokenize(formula):
        if token.kind != "ref":
            continue
        _, reference = split_sheet(token.text)
        parts: list[str] = []
        whole = False
        for corner in reference.split(":"):
            a1 = _A1.fullmatch(corner)
            assert a1 is not None
            col_fixed, letters, row_fixed, digits = a1.groups()
            whole = whole or not (letters and digits)
            axes: list[str] = []
            for axis, value, fixed, origin in (("R", int(digits) if digits else None, row_fixed or (col_fixed if not letters else ""), row),
                                               ("C", column_number(letters) if letters else None, col_fixed, column)):
                if value is not None:
                    axes.append(axis + (str(value) if fixed else (f"[{value - origin}]" if value != origin else "")))
            parts.append("".join(axes))
        if whole and len(parts) == 2 and parts[0] == parts[1]:
            # A whole row or column is written once: A:A is C, not C:C.
            parts = parts[:1]
        pieces.append((token.at, token.at + len(token.text), _prefix(token.text, reference) + ":".join(parts)))
    for start, end, replacement in reversed(pieces):
        formula = formula[:start] + replacement + formula[end:]
    return formula
