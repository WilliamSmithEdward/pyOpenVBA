"""Breaking formula text into tokens.

The text is the formula as a file stores it: no leading ``=``, functions
newer than Excel 2007 behind ``_xlfn.``, and English syntax whatever the
locale, so commas between arguments and a full stop in a number.

The operands come out already built as tree nodes: a number, a string, a
reference with its sheet qualifier. Deciding what an operand *is* needs the
characters around it, which the lexer has and the parser does not:

- ``LOG10`` is a cell, and ``LOG10(`` a function
- ``1:3`` is three whole rows, never the number 1
- ``A:C`` is three whole columns, and ``Jan:Mar!A1`` a 3D reference
- ``A1#`` is the range that spills from A1, and ``#N/A`` an error
- ``Table1[Qty]`` is a column of a table

Whitespace is kept as a token, because a space between two references is
an operator: ``A1:C3 B2:D4`` is the cells the two share.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum

from pyopenvba.formula._calc.nodes import (
    SPECIAL_ITEMS,
    THIS_ROW,
    AreaReference,
    AxisReference,
    CellReference,
    ErrorLiteral,
    ErrorReference,
    Logical,
    NameReference,
    Node,
    Number,
    Prefix,
    StructuredReference,
    Text,
)
from pyopenvba.formula._calc.numbers import literal_value
from pyopenvba.formula._calc.reference import MAX_COLUMN, MAX_ROW, AxisRef, CellRef, column_index


class FormulaSyntaxError(ValueError):
    """Formula text Excel would not accept, with where it goes wrong."""

    def __init__(self, message: str, formula: str, position: int) -> None:
        super().__init__(f"{message} at position {position} of {formula!r}")
        self.formula = formula
        self.position = position


class Kind(Enum):
    OPERAND = "operand"
    #: A name and the ``(`` after it.
    FUNCTION = "function"
    OPERATOR = "operator"
    COMMA = "comma"
    SEMICOLON = "semicolon"
    OPEN = "open"
    CLOSE = "close"
    OPEN_ARRAY = "open_array"
    CLOSE_ARRAY = "close_array"
    SPACE = "space"
    END = "end"


@dataclass(frozen=True, slots=True)
class Token:
    kind: Kind
    text: str
    start: int
    #: The operand, for an ``OPERAND``.
    node: Node | None = None


#: The errors a formula can name, longest first so ``#N/A`` does not stop
#: ``#NAME?`` short. Excel writes them in upper case and reads any case.
ERROR_CODES = (
    "#GETTING_DATA",
    "#CONNECT!",
    "#BLOCKED!",
    "#UNKNOWN!",
    "#PYTHON!",
    "#EXTERNAL!",
    "#DIV/0!",
    "#VALUE!",
    "#FIELD!",
    "#SPILL!",
    "#CALC!",
    "#BUSY!",
    "#NULL!",
    "#NAME?",
    "#NUM!",
    "#REF!",
    "#N/A",
)

_SPACE = re.compile(r"[ \t\r\n]+")
_NUMBER = re.compile(r"(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?")
#: What may not follow a reference, because the reference would then be the
#: start of something longer: a name, a call, a table.
_END = r"(?![\w.\\?$(\[])"
_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})" + _END)
_AREA = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7}):(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})" + _END)
_COLUMNS = re.compile(r"(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})" + _END)
_ROWS = re.compile(r"(\$?)([0-9]{1,7}):(\$?)([0-9]{1,7})" + _END)
#: A defined name or a function: a letter, underscore or backslash, then
#: letters, digits, full stops, underscores, backslashes and question marks.
_NAME = re.compile(r"(?:[^\W\d]|\\)[\w.\\?]*")
#: A sheet name that needs no quotes, and a 3D pair of them, before ``!``.
_BARE_SHEETS = re.compile(r"((?:[^\W\d]|\\)[\w.\\]*)(?::((?:[^\W\d]|\\)[\w.\\]*))?!")
_BOOK = re.compile(r"\[([^\]\[]+)\]")
_TWO_CHARACTER_OPERATORS = ("<>", "<=", ">=")
_OPERATORS = "+-*/^&=<>%:@"
#: "[" is not here: it starts an operand, a table's column or a workbook.
_PUNCTUATION = {
    "(": Kind.OPEN,
    ")": Kind.CLOSE,
    "{": Kind.OPEN_ARRAY,
    "}": Kind.CLOSE_ARRAY,
    ",": Kind.COMMA,
    ";": Kind.SEMICOLON,
}


def tokenize(formula: str) -> list[Token]:
    """The tokens of ``formula``, ending with an ``END`` token."""
    return _Lexer(formula).run()


class _Lexer:
    def __init__(self, formula: str) -> None:
        self.text = formula
        self.position = 0
        self.tokens: list[Token] = []

    def fail(self, message: str, position: int | None = None) -> FormulaSyntaxError:
        return FormulaSyntaxError(message, self.text, self.position if position is None else position)

    def emit(self, kind: Kind, end: int, node: Node | None = None) -> None:
        self.tokens.append(Token(kind, self.text[self.position : end], self.position, node))
        self.position = end

    def run(self) -> list[Token]:
        text = self.text
        while self.position < len(text):
            char = text[self.position]
            space = _SPACE.match(text, self.position)
            if space:
                self.emit(Kind.SPACE, space.end())
            elif char == '"':
                self.string()
            elif char == "#":
                self.hash()
            elif char in _PUNCTUATION:
                self.emit(_PUNCTUATION[char], self.position + 1)
            elif char in "'[$." or char.isdigit() or _NAME.match(text, self.position):
                self.operand()
            elif text.startswith(_TWO_CHARACTER_OPERATORS, self.position):
                self.emit(Kind.OPERATOR, self.position + 2)
            elif char in _OPERATORS:
                self.emit(Kind.OPERATOR, self.position + 1)
            else:
                raise self.fail(f"unexpected {char!r}")
        self.tokens.append(Token(Kind.END, "", len(text)))
        return self.tokens

    def string(self) -> None:
        end = _end_of_quoted(self.text, self.position, '"')
        if end is None:
            raise self.fail("a string is not closed")
        raw = self.text[self.position + 1 : end - 1]
        self.emit(Kind.OPERAND, end, Text(raw.replace('""', '"')))

    def hash(self) -> None:
        """``#`` after an operand is the spill operator; elsewhere it
        starts an error."""
        previous = self.tokens[-1] if self.tokens else None
        if previous is not None and previous.kind in (Kind.OPERAND, Kind.CLOSE):
            self.emit(Kind.OPERATOR, self.position + 1)
            return
        upper = self.text[self.position : self.position + 14].upper()
        for code in ERROR_CODES:
            if upper.startswith(code):
                self.emit(Kind.OPERAND, self.position + len(code), ErrorLiteral(code))
                return
        raise self.fail("not an error Excel has")

    def operand(self) -> None:
        text = self.text
        start = self.position
        prefix, after = self.prefix()
        if prefix is not None:
            node, end = self.target(after, prefix)
            if node is None:
                raise self.fail("nothing to refer to after '!'", after)
            self.emit(Kind.OPERAND, end, node)
            return

        char = text[start]
        if char == "[":
            end = self.brackets(start)
            self.emit(Kind.OPERAND, end, self.structured(None, text[start + 1 : end - 1]))
            return
        if char == "'":
            # A name that looks like a cell, which R1C1 read as a name: 'A1' (tests/fixtures/formula_notation.json).
            quoted = self.quoted_name(start, None)
            if quoted is None:
                raise self.fail("a quoted name is not closed")
            self.emit(Kind.OPERAND, quoted[1], quoted[0])
            return

        node, end = self.target(start, None)
        if node is not None:
            self.emit(Kind.OPERAND, end, node)
            return

        number = _NUMBER.match(text, start)
        if number and (char.isdigit() or char == "."):
            digits = number.group(0)
            # Fifteen significant digits, the rest cut, as Excel reads a
            # literal whether typed or loaded from a file.
            self.emit(Kind.OPERAND, number.end(), Number(literal_value(digits), digits))
            return

        name = _NAME.match(text, start)
        if name is None:
            raise self.fail(f"unexpected {char!r}")
        end = name.end()
        word = name.group(0)
        if end < len(text) and text[end] == "(":
            self.emit(Kind.FUNCTION, end + 1)
            return
        if end < len(text) and text[end] == "[":
            close = self.brackets(end)
            self.emit(Kind.OPERAND, close, self.structured(word, text[end + 1 : close - 1]))
            return
        if word.upper() in ("TRUE", "FALSE"):
            self.emit(Kind.OPERAND, end, Logical(word.upper() == "TRUE"))
            return
        self.emit(Kind.OPERAND, end, NameReference(word))

    def prefix(self) -> tuple[Prefix | None, int]:
        """The sheet qualifier at the current position, if there is one,
        and where the reference after its ``!`` starts."""
        text = self.text
        start = self.position
        if text[start] == "'":
            end = _end_of_quoted(text, start, "'")
            if end is None or end >= len(text) or text[end] != "!":
                return None, start
            return _quoted_prefix(text[start + 1 : end - 1].replace("''", "'")), end + 1
        book: str | None = None
        position = start
        if text[start] == "[":
            found = _BOOK.match(text, start)
            if found is None:
                return None, start
            if found.end() < len(text) and text[found.end()] == "!":
                return Prefix(book=found.group(1)), found.end() + 1
            book = found.group(1)
            position = found.end()
        sheets = _BARE_SHEETS.match(text, position)
        if sheets is None:
            return None, start
        return Prefix(sheets.group(1), sheets.group(2), book), sheets.end()

    def target(self, start: int, prefix: Prefix | None) -> tuple[Node | None, int]:
        """The reference at ``start``: a cell, an area, whole rows or
        columns, ``#REF!``, or with a prefix a name. ``None`` when there is
        none, which without a prefix leaves the text to be a number, a
        name or a call."""
        text = self.text
        if text.startswith("#REF!", start):
            return ErrorReference(prefix), start + 5
        rows = _ROWS.match(text, start)
        if rows:
            axis = _axis(rows, is_row=True)
            if axis is not None:
                return AxisReference(axis, prefix), rows.end()
        area = _AREA.match(text, start)
        if area:
            first = _cell(area.group(1), area.group(2), area.group(3), area.group(4))
            last = _cell(area.group(5), area.group(6), area.group(7), area.group(8))
            if first is not None and last is not None:
                return AreaReference(first, last, prefix), area.end()
        cell = _CELL.match(text, start)
        if cell:
            ref = _cell(*cell.groups())
            if ref is not None:
                return CellReference(ref, prefix), cell.end()
        columns = _COLUMNS.match(text, start)
        if columns:
            axis = _axis(columns, is_row=False)
            if axis is not None:
                return AxisReference(axis, prefix), columns.end()
        if prefix is None:
            return None, start
        if text[start : start + 1] == "'":
            quoted = self.quoted_name(start, prefix)
            return (None, start) if quoted is None else quoted
        name = _NAME.match(text, start)
        if name is None:
            return None, start
        return NameReference(name.group(0), prefix), name.end()

    def quoted_name(self, start: int, prefix: Prefix | None) -> tuple[NameReference, int] | None:
        """The name in quotes at ``start``, 'A1', and where it ends; None if the quotes are not closed."""
        end = _end_of_quoted(self.text, start, "'")
        if end is None:
            return None
        return NameReference(self.text[start + 1 : end - 1].replace("''", "'"), prefix), end

    def structured(self, table: str | None, content: str) -> StructuredReference:
        try:
            return structured(table, content)
        except ValueError as error:
            raise self.fail(str(error)) from error

    def brackets(self, start: int) -> int:
        """Just past the ``]`` that closes the ``[`` at ``start``."""
        depth = 0
        index = start
        text = self.text
        while index < len(text):
            char = text[index]
            if char == "'":
                index += 2
                continue
            if char == "[":
                depth += 1
            elif char == "]":
                depth -= 1
                if depth == 0:
                    return index + 1
            index += 1
        raise self.fail("a '[' is not closed", start)


def _end_of_quoted(text: str, start: int, quote: str) -> int | None:
    """Just past the quote closing the run at ``start``, doubling being the
    escape, or ``None`` if nothing closes it."""
    index = start + 1
    while index < len(text):
        if text[index] == quote:
            if index + 1 < len(text) and text[index + 1] == quote:
                index += 2
                continue
            return index + 1
        index += 1
    return None


def _quoted_prefix(content: str) -> Prefix:
    """``'[1]Q1 Data'``'s content as a prefix. A sheet name cannot hold
    ``:``, so one inside the quotes separates the ends of a 3D range."""
    book: str | None = None
    found = _BOOK.match(content)
    if found:
        book = found.group(1)
        content = content[found.end() :]
    first, colon, last = content.partition(":")
    return Prefix(first, last if colon else None, book)


def _cell(dollar_column: str, letters: str, dollar_row: str, digits: str) -> CellRef | None:
    """A cell, or ``None`` past the last row or column Excel has."""
    row = int(digits)
    if not 1 <= row <= MAX_ROW or len(letters) > 3:
        return None
    try:
        column = column_index(letters)
    except ValueError:
        return None
    return CellRef(row, column, dollar_row == "$", dollar_column == "$")


def _axis(match: re.Match[str], *, is_row: bool) -> AxisRef | None:
    low_dollar, low, high_dollar, high = match.groups()
    try:
        if is_row:
            ends = [(int(low), low_dollar == "$"), (int(high), high_dollar == "$")]
            limit = MAX_ROW
        else:
            ends = [(column_index(low), low_dollar == "$"), (column_index(high), high_dollar == "$")]
            limit = MAX_COLUMN
    except ValueError:
        return None
    if not all(1 <= value <= limit for value, _ in ends):
        return None
    ends.sort(key=lambda end: end[0])
    return AxisRef(is_row, ends[0][0], ends[1][0], ends[0][1], ends[1][1])


# ----------------------------------------------------------------------
# Structured references
# ----------------------------------------------------------------------

_ITEM_SPELLING = {item.upper(): item for item in SPECIAL_ITEMS}


def structured(table: str | None, content: str) -> StructuredReference:
    """A structured reference from what stands between its outer brackets.

    The forms, for a table ``T``: ``T[]`` the data, ``T[Qty]`` a column,
    ``T[#All]`` a special item, ``T[[#Headers],[Qty]:[Price]]`` items and
    columns together, and ``T[@Qty]``, the form typed for this row, which a
    file spells ``T[[#This Row],[Qty]]``.
    """
    body = content.strip()
    first: str | None = None
    last: str | None = None
    if not body:
        return StructuredReference(table)
    if body.startswith("@"):
        rest = body[1:].strip()
        if rest.startswith("["):
            first, last = _columns(rest)
        elif rest:
            first = last = _unescape(rest)
        return StructuredReference(table, (THIS_ROW,), first, last)
    if body.startswith("#"):
        return StructuredReference(table, (_item(body),))
    if not body.startswith("["):
        first = last = _unescape(body)
        return StructuredReference(table, (), first, last)
    items: list[str] = []
    for part in _split_parts(body):
        if part.startswith("[#"):
            items.append(_item(part[1:-1]))
        else:
            first, last = _columns(part)
    return StructuredReference(table, tuple(items), first, last)


def _item(text: str) -> str:
    spelled = _ITEM_SPELLING.get(" ".join(text.split()).upper())
    if spelled is None:
        raise ValueError(f"{text!r} is not a special item of a table")
    return spelled


def _columns(part: str) -> tuple[str, str]:
    """``[Qty]`` or ``[Qty]:[Price]``, as its first and last column."""
    pieces = _split_parts(part, separator=":")
    names = [_unescape(piece.strip()[1:-1]) for piece in pieces]
    return names[0], names[-1]


def _split_parts(body: str, separator: str = ",") -> list[str]:
    """``[a],[b]:[c]`` split at the separators outside brackets."""
    parts: list[str] = []
    depth = 0
    current: list[str] = []
    index = 0
    while index < len(body):
        char = body[index]
        if char == "'" and index + 1 < len(body):
            current.append(body[index : index + 2])
            index += 2
            continue
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
        if char == separator and depth == 0:
            parts.append("".join(current).strip())
            current = []
        else:
            current.append(char)
        index += 1
    parts.append("".join(current).strip())
    return [part for part in parts if part]


def _unescape(name: str) -> str:
    """A column name with its ``'`` escapes resolved."""
    out: list[str] = []
    index = 0
    while index < len(name):
        if name[index] == "'" and index + 1 < len(name):
            out.append(name[index + 1])
            index += 2
            continue
        out.append(name[index])
        index += 1
    return "".join(out)


__all__ = ["ERROR_CODES", "FormulaSyntaxError", "Kind", "Token", "structured", "tokenize"]
