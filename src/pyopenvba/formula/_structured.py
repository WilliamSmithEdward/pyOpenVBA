"""Structured references: how a formula spells part of a table, on screen and in a file, and the cells it names.

Measured in live Excel (scripts/measure_structured_references.py, tests/fixtures/structured_references/):

* Range.Formula spells a reference in the table's and the columns' own
  case, the special items as #All, #Data, #Headers and #Totals, and this
  row as @: Table1[@Qty], Table1[@[Qty]:[Price]], Table1[@]. A file
  spells this row [#This Row], Table1[[#This Row],[Qty]], and the whole
  table Table1[] where Range.Formula has Table1.
* In a column's name an apostrophe escapes [ ] # and ', and @ on screen,
  where it would otherwise mean this row. After @ a name holding a space,
  a control character or punctuation other than | is written in brackets,
  Odd[@[Unit Price]]; a name with a space at either end is in brackets on
  its own too, Chars[[ ab]]. Nothing else is: Odd[a,b] and Odd[Total $].
* Excel keeps a space inside each bracket and one after each comma, or
  none: Table1[ [#All] , [Qty] ] reads back Table1[ [#All], [Qty] ].
* A formula inside a table leaves the table's name out of a reference to
  one of its columns, [Qty] or [@Qty], and keeps it in any other.
* Where one value is wanted, a column of a table with more than one row
  of data is read as this row's cell, and written so: =Table1[Qty] reads
  back =Table1[@Qty]. A column of a one-row table stays as it is.

The model keeps a formula as Range.Formula spells it, but with every
reference naming its table; :func:`shown` leaves the name out where
Excel does, and :func:`in_file` and :func:`from_file` turn it into the
file's spelling and back.
"""

from __future__ import annotations

from collections.abc import Callable, Collection
from dataclasses import dataclass, replace
from typing import Final

from pyopenvba._a1 import Area
from pyopenvba.formula._parse import (ALL, DATA, HEADERS, THIS_ROW, TOTALS, FormulaError, Structured, Token,
                                      read_structured, tokenize)
from pyopenvba.formula._values import REF, VALUE

#: Characters that put a column's name in brackets after @: control characters, the space and ASCII punctuation
#: except | (the character sweep in tests/fixtures/structured_references/structured.json).
_BRACKETED_AFTER_AT: Final = (frozenset(chr(code) for code in range(0x21))
                              | frozenset("!\"#$%&'()*+,-./:;<=>?@[\\]^_`{}~"))


@dataclass(frozen=True, slots=True)
class TableShape:
    """What a structured reference needs to know of a table."""

    name: str
    #: The whole table, header and totals rows included, on its sheet.
    area: Area
    headers: bool
    totals: bool
    columns: tuple[str, ...]

    @property
    def data_rows(self) -> int:
        return self.area.rows - self.headers - self.totals

    def column(self, name: str) -> int | None:
        """A column's place in the table, counted from 0, found in any case."""
        wanted = name.lower()
        return next((index for index, one in enumerate(self.columns) if one.lower() == wanted), None)


# --- spelling ------------------------------------------------------------------------------------


def spelled(node: Structured, *, file: bool = False, table: str | None = None, at_escaped: bool = False) -> str:
    """A reference as Range.Formula spells it, or as a file does with ``file``; ``table`` is the name written, and
    ``at_escaped`` keeps the ' before an @ in a column's name, which a file leaves bare."""
    name = node.table if table is None else table
    file_escapes = file and not at_escaped
    items = node.items
    if not items and node.first is None and not node.spaced:
        return name + "[]" if file else name
    if not file and items == (THIS_ROW,) and not node.spaced:
        if node.first is None:
            return name + "[@]"
        if node.span:
            return f"{name}[@[{_escaped(node.first, file)}]:[{_escaped(node.last or node.first, file)}]]"
        escaped = _escaped(node.first, file)
        return f"{name}[@[{escaped}]]" if any(char in _BRACKETED_AFTER_AT for char in node.first) \
            else f"{name}[@{escaped}]"
    if not items and node.first is not None and not node.span and not node.spaced:
        escaped = _escaped(node.first, file_escapes)
        return f"{name}[[{escaped}]]" if node.first != node.first.strip(" ") else f"{name}[{escaped}]"
    if len(items) == 1 and node.first is None and not node.spaced:
        return f"{name}[{items[0]}]"
    parts = [f"[{item}]" for item in items]
    if node.first is not None:
        parts.append(f"[{_escaped(node.first, file_escapes)}]"
                     + (f":[{_escaped(node.last or node.first, file_escapes)}]" if node.span else ""))
    inside = (", " if node.comma_spaced else ",").join(parts)
    return f"{name}[ {inside} ]" if node.spaced else f"{name}[{inside}]"


def _escaped(name: str, file: bool) -> str:
    special = "[]#'" if file else "[]#'@"
    return "".join("'" + char if char in special else char for char in name)


def own_column(node: Structured) -> bool:
    """Whether a formula inside the table leaves the table's name out of this reference: one column, or this
    row's cell of one."""
    return node.first is not None and not node.span and node.items in ((), (THIS_ROW,))


def one_cell(node: Structured) -> Structured:
    """A column where one value is wanted, as Excel writes it: this row's cell of it."""
    return replace(node, items=(THIS_ROW,))


# --- rewriting a formula -----------------------------------------------------------------------


def rewritten(formula: str, change: Callable[[Token], str | None]) -> str:
    """``formula`` with each token ``change`` answers for replaced by the answer, everything else as written."""
    body = formula[1:] if formula.startswith("=") else formula
    try:
        tokens = tokenize(body)
    except FormulaError:
        return formula
    pieces: list[tuple[int, int, str]] = []
    for token in tokens:
        replacement = change(token)
        if replacement is not None and replacement != token.text:
            pieces.append((token.at, token.at + len(token.text), replacement))
    if not pieces:
        return formula
    out = body
    for start, stop, replacement in reversed(pieces):
        out = out[:start] + replacement + out[stop:]
    return ("=" if formula.startswith("=") else "") + out


def has_tables(formula: str) -> bool:
    """Whether a formula might name a table: a quick look before tokenizing it."""
    return "[" in formula


def shown(formula: str, here: str | None) -> str:
    """A formula as Range.Formula shows it in a cell of the table named ``here``, or of none."""
    if here is None or "[" not in formula:
        return formula
    wanted = here.lower()

    def change(token: Token) -> str | None:
        if token.kind != "structured":
            return None
        node = read_structured(token.text)
        if node.table.lower() != wanted or not own_column(node):
            return None
        return spelled(node, table="")

    return rewritten(formula, change)


def in_file(formula: str, tables: Collection[str], *, at_escaped: bool = False) -> str:
    """A formula as a file spells it: this row as [#This Row], a table on its own as Table1[], @ unescaped; with
    ``at_escaped``, as the formula engine reads it, a column's @ kept escaped, since the engine reads a bare [@ as
    this row."""
    wanted = {name.lower() for name in tables}
    lowered = formula.lower()
    if "[" not in formula and not any(name in lowered for name in wanted):
        return formula

    def change(token: Token) -> str | None:
        if token.kind == "structured":
            return spelled(read_structured(token.text), file=True, at_escaped=at_escaped)
        if token.kind == "name" and token.text.lower() in wanted:
            return token.text + "[]"
        return None

    return rewritten(formula, change)


def from_file(formula: str) -> str:
    """A formula a file spells, as Range.Formula spells it with every table named."""
    if "[" not in formula:
        return formula

    def change(token: Token) -> str | None:
        return spelled(read_structured(token.text, file=True)) if token.kind == "structured" else None

    return rewritten(formula, change)


def renamed_table(formula: str, old: str, new: str) -> str:
    """A formula after the table named ``old`` is renamed ``new``: in its references and where it stands alone."""
    wanted = old.lower()
    if wanted not in formula.lower():
        return formula

    def change(token: Token) -> str | None:
        if token.kind == "structured":
            node = read_structured(token.text)
            return spelled(replace(node, table=new)) if node.table.lower() == wanted else None
        if token.kind == "name" and token.text.lower() == wanted:
            return new
        return None

    return rewritten(formula, change)


def renamed_columns(formula: str, table: str, names: dict[str, str]) -> str:
    """A formula after columns of ``table`` are renamed, ``names`` taking each old name, in lower case, to its new."""
    wanted = table.lower()
    if "[" not in formula:
        return formula

    def change(token: Token) -> str | None:
        if token.kind != "structured":
            return None
        node = read_structured(token.text)
        if node.table.lower() != wanted or node.first is None:
            return None
        first = names.get(node.first.lower(), node.first)
        last = names.get((node.last or node.first).lower(), node.last or node.first)
        return spelled(replace(node, first=first, last=last))

    return rewritten(formula, change)


def filled(formula: str, across: int, columns: Callable[[str], tuple[str, ...] | None]) -> str:
    """A formula AutoFill carries ``across`` columns: a reference to one column of a table moves along the
    table's columns as far, round from the last to the first; a range of columns, a whole table or a special
    item alone stays (tests/fixtures/structured_references/structured.json)."""
    if not across or "[" not in formula:
        return formula

    def change(token: Token) -> str | None:
        if token.kind != "structured":
            return None
        node = read_structured(token.text)
        names = columns(node.table)
        if node.first is None or node.span or not names:
            return None
        lowered = [name.lower() for name in names]
        if node.first.lower() not in lowered:
            return None
        moved = names[(lowered.index(node.first.lower()) + across) % len(names)]
        return spelled(replace(node, first=moved, last=moved))

    return rewritten(formula, change)


# --- what a reference names ------------------------------------------------------------------


def area(node: Structured, table: TableShape, row: int) -> Area:
    """The cells a reference names, for a formula on ``row``.

    A special item the table does not show is #REF! -- the totals row of
    a table without one -- except beside the data, where the data is
    what is left: Table1[[#Data],[#Totals]] is its data. This row is the
    formula's own row, on whatever sheet the formula is, and #VALUE! when
    that row is not one of the table's data rows. A column the table has
    not got is #REF!.
    """
    whole = table.area
    data_top = whole.top + table.headers
    data_bottom = whole.bottom - table.totals
    items = set(node.items) or {DATA}
    if items == {ALL}:
        top, bottom = whole.top, whole.bottom
    elif items == {THIS_ROW}:
        if not data_top <= row <= data_bottom:
            raise VALUE
        top = bottom = row
    else:
        spans: list[tuple[int, int]] = []
        if HEADERS in items and table.headers:
            spans.append((whole.top, whole.top))
        if DATA in items:
            spans.append((data_top, data_bottom))
        if TOTALS in items and table.totals:
            spans.append((whole.bottom, whole.bottom))
        if not spans:
            raise REF
        top, bottom = min(span[0] for span in spans), max(span[1] for span in spans)
    left, right = whole.left, whole.right
    if node.first is not None:
        first = table.column(node.first)
        last = table.column(node.last or node.first)
        if first is None or last is None:
            raise REF
        left, right = whole.left + min(first, last), whole.left + max(first, last)
    return Area(top, left, bottom, right, whole.sheet)
