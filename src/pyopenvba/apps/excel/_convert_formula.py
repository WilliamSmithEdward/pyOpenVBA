"""Application.ConvertFormula: a formula's references turned from A1 to R1C1 or back, and made absolute or relative,
as Excel converts them (scripts/measure_convert_formula.py, tests/fixtures/convert_formula.json).

Excel reads the formula as a formula on no workbook's sheet:

- A reference to a sheet of the active workbook names the workbook as well, ``[Book1]Data!A1``, in apostrophes
  where the workbook's or the sheet's name needs them, and a name with a sheet before it is that sheet's own,
  ``[Book1]Data!Local``, or else the workbook's, ``Book1!Name``. A sheet the workbook has not got keeps its
  spelling, as another workbook's sheet does, and the active workbook's own names are never looked up.
- Every other name, and a function Excel has not got, is spelled as Excel first saw it in any formula it has
  converted, whatever case it comes in after; a name a LET binds as the formula first writes it.
- Relative references in R1C1, given and written, are relative to RelativeTo's first cell, or else the active
  cell; an offset wraps round the sheet's edge as it is read, R[-1]C in the first row being the last row, and is
  written as the difference, R[1048575]C.
- Spaces stay where they are, except at the end, before a comma between a function's arguments, and between two
  cells a colon joins, which A1 :B2 is.

What Excel cannot read it answers with #VALUE!, the error value 2015: a table's reference, a space in an array
constant, a name in apostrophes, a tab or a carriage return, an offset as long as the sheet, a formula of more
than 255 characters, and one whose text does not fit as :class:`_Buffer` writes it.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import TYPE_CHECKING, Final

from pyopenvba._a1 import column_letter, column_number, parse_area, quote_sheet, quote_workbook
from pyopenvba._a1 import split_sheet as split_name
from pyopenvba.apps.excel import _names
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.catalog import is_excel_function
from pyopenvba.formula._parse import FormulaError, Node, Token, literal, parse, tokenize
from pyopenvba.formula._prefixes import bound
from pyopenvba.formula._r1c1 import is_cell, refused_in_a1, to_a1
from pyopenvba.formula._spell import LARGEST, check_arguments
from pyopenvba.formula._values import number_text
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import (EMPTY, MISSING, NOTHING, NULL, VBAArray, VBAErrorValue, error,
                                           to_integer, to_number, to_text)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Workbook

#: #VALUE!, what ConvertFormula answers for a formula it cannot convert.
_VALUE: Final = 2015
#: The longest formula ConvertFormula takes, in characters.
_LONGEST: Final = 255
#: How much the text Excel writes holds; each piece of it counts one character more while it is being written.
_ROOM: Final = 254
#: The room Excel wants before it writes a whole number up to 65535; any other number, or a reference past its
#: sheet; a number in an array constant; past a function's name; and after a function Excel has not got.
_WHOLE: Final = 5
_OPERAND: Final = 25
_ITEM: Final = 22
_PAST_NAME: Final = 4
_UNKNOWN_CALL: Final = 2
#: Loosest first, as pyopenvba.formula._parse reads them; ^ binds tighter, and the signs tighter still.
_LEVELS: Final = (("=", "<>", "<", ">", "<=", ">="), ("&",), ("+", "-"), ("*", "/"))
#: Which coordinates each ToAbsolute fixes: xlAbsolute, xlAbsRowRelColumn, xlRelRowAbsColumn and xlRelative,
#: the row's first.
_FIXED: Final = {1: (True, True), 2: (True, False), 3: (False, True), 4: (False, False)}
_CORNER: Final = re.compile(r"(\$?)([A-Za-z]{1,3})?(\$?)([0-9]{1,7})?")
#: RelativeTo given as text: an R1C1 reference, with an = and a sheet before it or not.
_RELATIVE_TO: Final = re.compile(r"=?(?:(?:'(?:[^']|'')+'|[^'!]+)!)?(?P<area>[^!]*)")
_R1C1_AREA: Final = re.compile(
    r"(?:R(?:\[-?\d+\]|\d+)?)?(?:C(?:\[-?\d+\]|\d+)?)?(?::(?:R(?:\[-?\d+\]|\d+)?)?(?:C(?:\[-?\d+\]|\d+)?)?)?",
    re.IGNORECASE)


class _Refused(Exception):
    """A formula, or an argument, that ConvertFormula answers #VALUE! for."""


#: A reference's corner: its column and its row, each fixed or not with its number, None in a whole row or column.
_Corner = tuple[tuple[bool, int] | None, tuple[bool, int] | None]


def convert_formula(application: Application, formula: object, start: object, end: object, absolute: object,
                    relative: object) -> object:
    """Application.ConvertFormula(Formula, FromReferenceStyle, ToReferenceStyle, ToAbsolute, RelativeTo).

    FromReferenceStyle is a Long, xlR1C1 and 0 meaning R1C1 and any other number A1. A range given as the formula
    has each cell's value converted, answered as an array in the range's shape, a one-dimensional one for a single
    row; an error value is answered as it is, and Empty and Null with #VALUE!.
    """
    from pyopenvba.apps.excel._model import Range

    if formula is MISSING or start is MISSING:
        raise error(449)
    if start is NULL:
        raise error(94)
    given = int(to_integer(start, "Long")) in (0, -4150)
    try:
        wanted: bool | VBAErrorValue = _style(end, given)
        mode = _absolute(absolute)
        origin = _origin(application, relative)
    except _Refused:
        wanted, mode, origin = VBAErrorValue(_VALUE), None, (1, 1)

    def one(value: object) -> object:
        if isinstance(value, VBAErrorValue):
            return value
        if value is NULL or value is EMPTY:
            return VBAErrorValue(_VALUE)
        if isinstance(value, (VBAArray, VBAObject)):
            raise VBAUnsupportedError("Application.ConvertFormula given an array or an object is not implemented")
        if isinstance(wanted, VBAErrorValue):
            return wanted
        try:
            return _Converter(application, given, wanted, mode, origin).converted(to_text(value))
        except _Refused:
            return VBAErrorValue(_VALUE)

    if isinstance(formula, Range):
        return _cell_by_cell(formula.vba_get("Value"), one)
    return one(formula)


def _cell_by_cell(values: object, convert: Callable[[object], object]) -> object:
    """Each of a range's values converted: one row of them as a one-dimensional array, as Excel answers."""
    if not isinstance(values, VBAArray):
        return convert(values)
    items = [convert(item) for item in values.items]
    (_, rows), (_, columns) = values.bounds
    if rows == 1:
        return VBAArray([(1, columns)], items=items)
    return VBAArray([(1, rows), (1, columns)], items=items)


def _style(value: object, given: bool) -> bool | VBAErrorValue:
    """Whether ToReferenceStyle asks for R1C1: as FromReferenceStyle where it is left out or Null; xlR1C1, 0, False
    and Empty for R1C1, True and any other number for A1; the text TRUE or FALSE as that logical, any other text
    refused. An error value is the answer."""
    if value is MISSING or value is NULL:
        return given
    if isinstance(value, VBAErrorValue):
        return value
    if value is EMPTY:
        return True
    if isinstance(value, bool):
        return not value
    if isinstance(value, str):
        if value.upper() not in ("TRUE", "FALSE"):
            raise _Refused
        return value.upper() == "FALSE"
    return float(to_number(value)) in (0.0, -4150.0)


def _absolute(value: object) -> int | None:
    """ToAbsolute as Excel takes it: rounded down, text read as a number and True as 1; 0, False, Empty and Null
    keep each reference as it is, 1 to 4 make it so, and anything else is refused."""
    if value is MISSING or value is NULL or value is EMPTY:
        return None
    if isinstance(value, bool):
        number = 1.0 if value else 0.0
    elif isinstance(value, str):
        try:
            number = float(value)
        except ValueError:
            raise _Refused from None
    elif isinstance(value, VBAErrorValue):
        raise _Refused
    else:
        number = float(to_number(value))
    whole = math.floor(number) if math.isfinite(number) else -1
    if whole == 0:
        return None
    if whole not in _FIXED:
        raise _Refused
    return whole


def _origin(application: Application, value: object) -> tuple[int, int]:
    """The cell relative references are relative to: RelativeTo's first, an R1C1 reference given as text read at
    the active cell, or the active cell itself."""
    from pyopenvba.apps.excel._model import Range

    if isinstance(value, Range):
        return value.first.top, value.first.left
    if value is NOTHING:
        raise error(5)
    active = application.vba_get("ActiveCell")
    if not isinstance(active, Range):
        raise VBAUnsupportedError("Application.ConvertFormula with no active cell is not implemented")
    here = active.first.top, active.first.left
    if value is MISSING:
        return here
    if not isinstance(value, str):
        raise _Refused
    found = _RELATIVE_TO.fullmatch(value)
    if found is None or not found.group("area") or not _R1C1_AREA.fullmatch(found.group("area")):
        raise _Refused
    try:
        reference = to_a1(found.group("area"), *here, strict=True)
    except ValueError:
        raise _Refused from None
    if reference == found.group("area"):
        # R0C1 is a name, and no reference.
        raise _Refused
    area = parse_area(reference.replace("$", ""))
    return area.top, area.left


@dataclass(slots=True)
class _Piece:
    """One token as ConvertFormula writes it, the spaces before it apart; ``room`` is what Excel wants free before
    writing it past its spaces, ``known`` for a call of a function Excel has."""

    kind: str
    text: str
    spaces: str = ""
    room: int = 0
    known: bool = True


class _Converter:
    """One formula, read in one reference style and written in the other against the active workbook."""

    def __init__(self, application: Application, given: bool, wanted: bool, mode: int | None,
                 origin: tuple[int, int]) -> None:
        self.application = application
        self.book: Workbook | None = application.active_book
        self.given = given
        self.wanted = wanted
        self.mode = mode
        self.origin = origin

    def converted(self, text: str) -> str:
        if len(text) > _LONGEST:
            raise _Refused
        equals = text.startswith("=")
        body = (text[1:] if equals else text).rstrip(" \n")
        if not body:
            raise _Refused
        if self.given:
            try:
                body = to_a1(body, *self.origin, names=True, strict=True)
            except ValueError:
                raise _Refused from None
        pieces = self.pieces(self.tokens(body), body)
        written = "".join(piece.spaces + piece.text for piece in pieces)
        _Buffer(pieces).fill()
        return ("=" if equals else "") + written

    def tokens(self, body: str) -> list[Token]:
        """The formula's tokens, refused where Excel would not read the formula, and two cells a colon and spaces
        join made one."""
        try:
            tokens = tokenize(body, spaces=True)
        except FormulaError:
            raise _Refused from None
        braces = 0
        for token in tokens:
            braces += {"lbrace": 1, "rbrace": -1}.get(token.kind, 0)
            if token.kind == "structured" or token.kind == "ws" and (braces or token.text.strip(" \n")):
                raise _Refused
            if token.kind == "name" and token.text.rpartition("!")[2].startswith("'") and not self.given:
                raise _Refused
        bound_at = bound(body)
        if not self.given and any(token.kind == "name" and token.at in bound_at and is_cell(token.text)
                                  for token in tokens):
            # LET(x1,1,x1): A1 cannot read a cell as a name.
            raise _Refused
        try:
            check_arguments(_parsed(body, tokens))
        except FormulaError:
            raise _Refused from None
        return _joined_cells(tokens)

    def pieces(self, tokens: list[Token], body: str) -> list[_Piece]:
        """Each token as Excel writes it back."""
        bound_at = bound(body)
        declared: dict[str, str] = {}
        #: The brackets open, innermost last: True for a call's.
        calls: list[bool] = []
        out: list[_Piece] = []
        spaces = ""
        previous: Token | None = None
        index = 0
        while tokens[index].kind != "eof":
            token = tokens[index]
            if token.kind == "ws":
                if not (_next(tokens, index).kind == "comma" and calls[-1:] == [True]):
                    spaces += token.text
                index += 1
                continue
            piece = _Piece(token.kind, token.text, spaces)
            if token.kind == "open":
                # A call's bracket follows its name, or a LAMBDA's, (LAMBDA(x,x))(1), with no space between.
                calls.append(previous is not None and previous.kind in ("name", "close")
                             and previous.at + len(previous.text) == token.at)
            elif token.kind == "close" and calls:
                calls.pop()
            if token.kind == "op" and token.text in "+-" and _sign(previous) and tokens[index + 1].kind == "number":
                # A sign before a number, with no space between, is the number's own.
                index += 1
                piece.kind = "number"
                piece.text, piece.room = _number(tokens[index].text, negative=token.text == "-")
            elif token.kind == "number":
                piece.text, piece.room = _number(token.text, negative=False)
            elif token.kind == "error":
                head, _, name = token.text.rpartition("#")
                if head:
                    prefix = self.prefix(head[:-1])
                    piece.text, piece.room = prefix + "!#" + name.upper(), len(prefix) + 1 + _OPERAND
                else:
                    piece.text = token.text.upper()
            elif token.kind == "ref":
                piece.text, piece.room = self.reference(token.text)
            elif token.kind == "name" and token.at in bound_at:
                piece.text = declared.setdefault(token.text.upper(), token.text)
            elif token.kind == "name":
                following = tokens[index + 1]
                self.name(piece, following.kind == "open" and following.at == token.at + len(token.text))
            out.append(piece)
            spaces = ""
            previous = tokens[index]
            index += 1
        return out

    def name(self, piece: _Piece, called: bool) -> None:
        """A name, or a function's, as Excel spells it."""
        head, bang, bare = piece.text.rpartition("!")
        if called:
            piece.kind = "call"
            piece.known = not bang and is_excel_function(bare)
            piece.text = bare.upper() if piece.known else piece.text if bang else self.remembered(bare)
        elif bang:
            prefix, bare = self.qualified(head, bare)
            piece.text, piece.room = prefix + "!" + bare, len(prefix) + 1 + _OPERAND
        elif bare.upper() in ("TRUE", "FALSE") or is_excel_function(bare):
            # A logical, or a function's name with no bracket after it, which is a name, in capitals.
            piece.text = bare.upper()
        elif bare.startswith("'"):
            # A name that looks like a cell, which R1C1 read: in apostrophes only in A1.
            piece.text = bare[1:-1].replace("''", "'") if self.wanted else bare
        elif refused_in_a1(bare):
            raise _Refused
        else:
            piece.text = self.remembered(bare)

    def remembered(self, name: str) -> str:
        """The spelling Excel first saw a name in, which from now on is ``name``'s if it has none."""
        return self.application.convert_spellings.setdefault(name.lower(), name)

    def reference(self, text: str) -> tuple[str, int]:
        """A reference as Excel writes it, and the room it wants for it."""
        head, bang, body = text.rpartition("!")
        prefix = self.prefix(head) + "!" if bang else ""
        corners = [_corner(part) for part in body.split(":")]
        if len(corners) == 2:
            (first_column, first_row), (second_column, second_row) = corners
            first_column, second_column = _ordered(first_column, second_column)
            first_row, second_row = _ordered(first_row, second_row)
            corners = [(first_column, first_row), (second_column, second_row)]
        parts = [self.corner(column, row) for column, row in corners]
        whole = corners[0][0] is None or corners[0][1] is None
        if self.wanted and whole and len(parts) == 2 and parts[0] == parts[1]:
            # A whole row or column is written once in R1C1: A:A is C, not C:C.
            parts = parts[:1]
        return prefix + ":".join(parts), len(prefix) + _OPERAND

    def corner(self, column: tuple[bool, int] | None, row: tuple[bool, int] | None) -> str:
        if self.mode is not None:
            row_fixed, column_fixed = _FIXED[self.mode]
            column = None if column is None else (column_fixed, column[1])
            row = None if row is None else (row_fixed, row[1])
        if not self.wanted:
            return (("$" if column[0] else "") + column_letter(column[1]) if column else "") \
                + (("$" if row[0] else "") + str(row[1]) if row else "")
        top, left = self.origin
        return ("R" + (str(row[1]) if row[0] else _offset(row[1] - top)) if row else "") \
            + ("C" + (str(column[1]) if column[0] else _offset(column[1] - left)) if column else "")

    def prefix(self, head: str) -> str:
        """The sheet a reference names, as Excel writes it: one of the active workbook's, or a pair of them in its
        order, after the workbook's name; any other as it is written, apostrophes only where it needs them."""
        raw = head[1:-1].replace("''", "'") if head.startswith("'") and head.endswith("'") else head
        if raw.startswith("["):
            book, _, sheets = raw[1:].partition("]")
            if self.book is None or book.casefold() != self.book.name.casefold():
                return head
            own = self.own(sheets)
            if own is None:
                # The active workbook named outright, and a sheet it has not got.
                raise _Refused
            return own
        own = self.own(raw)
        if own is not None:
            return own
        return head if ":" in raw else quote_sheet(raw)

    def own(self, sheets: str) -> str | None:
        """A sheet of the active workbook, or a pair of them, as Excel names it, or None where it has not got
        them."""
        if self.book is None:
            return None
        names = [one.name for one in self.book.sheets_]
        folded = [one.casefold() for one in names]
        first, colon, last = sheets.partition(":")
        if first.casefold() not in folded or colon and last.casefold() not in folded:
            return None
        start, end = sorted((folded.index(first.casefold()), folded.index((last if colon else first).casefold())))
        return quote_workbook(self.book.name, names[start] if start == end else f"{names[start]}:{names[end]}")

    def qualified(self, head: str, bare: str) -> tuple[str, str]:
        """A name with a sheet before it: the sheet's own, the workbook's, or one nothing defines, on a sheet of the
        active workbook; any other kept as it is written."""
        raw = head[1:-1].replace("''", "'") if head.startswith("'") and head.endswith("'") else head
        book = self.book
        sheet = None if book is None or raw.startswith("[") or ":" in raw else self.own(raw)
        if book is None or sheet is None:
            return (head if raw.startswith("[") or ":" in raw else quote_sheet(raw)), bare
        own_name = next(one.name for one in book.sheets_ if one.name.casefold() == raw.casefold())
        local = _defined(book, own_name, bare)
        if local is not None:
            return sheet, local
        book_level = _defined(book, "", bare)
        if book_level is not None:
            return quote_workbook(book.name), book_level
        return sheet, bare


def _defined(book: Workbook, scope: str, bare: str) -> str | None:
    """A defined name's own spelling, on the sheet ``scope`` names or the workbook's where it is empty."""
    for entry in book.names_.entries:
        entry_scope, entry_bare = split_name(entry.name)
        if _names.same_name(entry_bare, bare) and _names.same_sheet_name(entry_scope, scope):
            return entry_bare
    return None


def _parsed(body: str, tokens: list[Token]) -> Node:
    """The formula read; a comma outside every bracket joins references, as a macro sheet's formula may."""
    try:
        return parse("=" + body)
    except FormulaError:
        depth = 0
        for token in tokens:
            depth += {"open": 1, "lbrace": 1, "close": -1, "rbrace": -1}.get(token.kind, 0)
            if token.kind == "comma" and not depth:
                return parse("=(" + body + ")")
        raise


def _joined_cells(tokens: list[Token]) -> list[Token]:
    """A1 :B2 as Excel reads it, one range: a cell, a colon and a cell with spaces about the colon. In
    A1 :B2:C3 the range is A1:B2, and C3 joins it after."""
    out: list[Token] = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        colon = _next_index(tokens, index)
        after = _next_index(tokens, colon) if tokens[colon].kind == "op" and tokens[colon].text == ":" else colon
        if token.kind == "ref" and is_cell(token.text.rpartition("!")[2]) and after != colon \
                and after - index > 2 and tokens[after].kind == "ref" and "!" not in tokens[after].text:
            first, _, rest = tokens[after].text.partition(":")
            if is_cell(first):
                out.append(Token("ref", token.text + ":" + first, token.at))
                if rest:
                    at = tokens[after].at + len(first)
                    out += [Token("op", ":", at), Token("ref", rest, at + 1)]
                index = after + 1
                continue
        out.append(token)
        index += 1
    return out


def _next_index(tokens: list[Token], index: int) -> int:
    """Where the first token after ``index`` that is not white space stands."""
    index += 1
    while index < len(tokens) - 1 and tokens[index].kind == "ws":
        index += 1
    return min(index, len(tokens) - 1)


def _next(tokens: list[Token], index: int) -> Token:
    return tokens[_next_index(tokens, index)]


def _sign(previous: Token | None) -> bool:
    """Whether a sign here stands before an operand rather than between two."""
    if previous is None:
        return True
    if previous.kind == "op":
        return previous.text not in ("%", "#")
    return previous.kind in ("open", "comma", "semicolon", "lbrace")


def _number(text: str, *, negative: bool) -> tuple[str, int]:
    """A number as Excel writes it back, and the room it wants: less for a whole one Excel keeps as an integer."""
    value = literal(text)
    if value > LARGEST:
        raise _Refused
    if negative and value:
        value = -value
    return number_text(value, formula=True), _WHOLE if value.is_integer() and 0 <= value <= 65535 else _OPERAND


def _corner(text: str) -> _Corner:
    """A corner as A1 writes it, $A1 or $1, read."""
    found = _CORNER.fullmatch(text)
    assert found is not None
    column_fixed, letters, row_fixed, digits = found.groups()
    column = (bool(column_fixed), column_number(letters)) if letters else None
    row = (bool(row_fixed or (column_fixed if not letters else "")), int(digits)) if digits else None
    return column, row


def _ordered(first: tuple[bool, int] | None, second: tuple[bool, int] | None) -> tuple[
        tuple[bool, int] | None, tuple[bool, int] | None]:
    if first is None or second is None or first[1] <= second[1]:
        return first, second
    return second, first


def _offset(difference: int) -> str:
    return f"[{difference}]" if difference else ""


class _Buffer:
    """The text Excel writes a converted formula into, as it writes it.

    It holds :data:`_ROOM` characters, and each piece of text it holds
    while it works counts one more. It writes each operand as it comes and
    joins operands under their operator once both are written, as Excel
    reads a formula's tokens; spaces go with what follows them. Before an
    operand it wants room for the longest the operand could be: five
    characters for a whole number up to 65535, 25 for any other number, and
    25 past its sheet for a reference or a name with a sheet before it; 22
    for each number in an array constant, past what of the array is written;
    four past a function's name, once its arguments are written; and two for
    a function Excel has not got, whose name is a piece of its own written
    first (scripts/measure_convert_formula.py, longest_* and chain_*).
    """

    def __init__(self, pieces: list[_Piece]) -> None:
        self.pieces = [*pieces, _Piece("eof", "")]
        self.at = 0
        self.used = 0

    def fill(self) -> None:
        whole = self.expression(0)
        while self.piece.kind == "comma":
            comma = self.take()
            whole = self.join([whole, self.expression(0)], len(comma.spaces + comma.text))
        assert self.piece.kind == "eof", f"the formula was read as far as {self.piece.text!r}"

    @property
    def piece(self) -> _Piece:
        return self.pieces[self.at]

    def take(self) -> _Piece:
        piece = self.pieces[self.at]
        self.at += 1
        return piece

    def is_op(self, *texts: str) -> bool:
        return self.piece.kind == "op" and self.piece.text in texts

    def fits(self, room: int) -> None:
        if self.used + room > _ROOM:
            raise _Refused

    def push(self, length: int) -> int:
        self.used += length + 1
        self.fits(0)
        return length

    def join(self, parts: list[int], added: int) -> int:
        """Pieces written joined into one, ``added`` characters of operators, brackets and names among them."""
        self.used -= sum(part + 1 for part in parts)
        return self.push(sum(parts) + added)

    def expression(self, level: int) -> int:
        if level == len(_LEVELS):
            return self.power()
        left = self.expression(level + 1)
        while self.is_op(*_LEVELS[level]):
            operator = self.take()
            left = self.join([left, self.expression(level + 1)], len(operator.spaces + operator.text))
        return left

    def power(self) -> int:
        node = self.tight()
        while self.is_op("^"):
            operator = self.take()
            node = self.join([node, self.tight()], len(operator.spaces + operator.text))
        return node

    def tight(self) -> int:
        if self.is_op("@"):
            operator = self.take()
            return self.percents(self.join([self.signed()], len(operator.spaces + operator.text)))
        if self.is_op("-", "+"):
            operator = self.take()
            return self.join([self.tight()], len(operator.spaces + operator.text))
        return self.percents(self.operand())

    def signed(self) -> int:
        if self.is_op("-", "+", "@"):
            operator = self.take()
            return self.join([self.signed()], len(operator.spaces + operator.text))
        return self.operand()

    def percents(self, node: int) -> int:
        while self.is_op("%"):
            operator = self.take()
            node = self.join([node], len(operator.spaces + operator.text))
        return node

    def operand(self) -> int:
        node = self.spilled(self.primary())
        while True:
            if self.is_op(":"):
                operator = self.take()
                node = self.join([node, self.primary()], len(operator.spaces + operator.text))
            elif self.piece.kind in ("ref", "name", "call", "open") and self.piece.spaces:
                # The spaces between two references are their intersection's operator.
                spaces, self.piece.spaces = self.piece.spaces, ""
                node = self.join([node, self.primary()], len(spaces))
            else:
                return node

    def spilled(self, node: int) -> int:
        while self.is_op("#"):
            operator = self.take()
            node = self.join([node], len(operator.spaces + operator.text))
        return node

    def primary(self) -> int:
        piece = self.piece
        if piece.kind == "open":
            self.take()
            inner = self.expression(0)
            while self.piece.kind == "comma":
                comma = self.take()
                inner = self.join([inner, self.expression(0)], len(comma.spaces + comma.text))
            close = self.take()
            return self.join([inner], len(piece.spaces + piece.text + close.spaces + close.text))
        if piece.kind == "lbrace":
            return self.array()
        if piece.kind == "call":
            node = self.call()
            while self.piece.kind == "open" and not self.piece.spaces:
                # A call of what a call gives, LAMBDA(x,x)(1), written as a function Excel has not got.
                node = self.arguments([node], "", _UNKNOWN_CALL)
            return node
        self.take()
        self.fits(len(piece.spaces) + piece.room)
        return self.push(len(piece.spaces + piece.text))

    def call(self) -> int:
        name = self.take()
        if name.known:
            return self.arguments([], name.spaces + name.text, len(name.text) + _PAST_NAME)
        return self.arguments([self.push(len(name.spaces + name.text))], "", _UNKNOWN_CALL)

    def arguments(self, parts: list[int], name: str, room: int) -> int:
        """A call's brackets and arguments, after ``parts`` written before them, and ``name`` written with them."""
        opening = self.take()
        added = len(name + opening.spaces + opening.text)
        if self.piece.kind != "close":
            while True:
                parts.append(self.push(0) if self.piece.kind in ("comma", "close") else self.expression(0))
                if self.piece.kind != "comma":
                    break
                comma = self.take()
                added += len(comma.spaces + comma.text)
        close = self.take()
        self.fits(room)
        return self.join(parts, added + len(close.spaces + close.text))

    def array(self) -> int:
        written = 0
        while True:
            piece = self.take()
            if piece.kind == "number":
                self.fits(written + len(piece.spaces) + _ITEM)
            written += len(piece.spaces + piece.text)
            if piece.kind == "rbrace":
                return self.push(written)
