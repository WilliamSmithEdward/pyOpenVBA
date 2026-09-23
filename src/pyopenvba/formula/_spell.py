"""How Excel spells a formula it is given.

Excel does not keep the text a formula is written with: it reads the
formula and writes it out again. Measured in live Excel
(scripts/measure_formula_spelling.py), through Range.Formula:

- A cell reference is in capitals, and a range runs from its top left
  corner to its bottom right, each coordinate keeping its own dollar
  sign: ``SUM($B$2:A1)`` becomes ``SUM(A1:$B$2)``, ``SUM(b:a)``
  ``SUM(A:B)``.
- A sheet is spelled as it is named, in apostrophes only where its name
  needs them. A sheet the workbook does not have keeps its spelling.
- A function Excel has is in capitals, and so are TRUE and FALSE and the
  errors. A defined name is spelled as it was defined. Any other name --
  a function Excel does not have, a macro's function, a name nothing
  defines -- is spelled as the workbook first saw it, whatever case it
  comes in after.
- A number keeps fifteen significant digits, the rest cut off, and is
  written out again as Range.Formula spells a stored one, a sign in front
  of it being part of it: ``=1.50`` is ``=1.5``, ``=-0`` is ``=0``, ``=+1``
  is ``=1`` and ``=123456789012345678`` is ``=123456789012345000``. One past
  9.99999999999999E+307 is refused.
- Spaces and line breaks stay where they are, except before a comma, at
  the end, and anywhere in an array constant.

What Excel refuses to read raises :class:`FormulaError`; what it reads
and the model does not -- a table's columns in brackets, the ``@`` and
``#`` of dynamic arrays -- raises :class:`UnmodelledFormulaError`.
"""

from __future__ import annotations

import re
from typing import Final, Protocol

from pyopenvba._a1 import column_letter, column_number, quote_sheet
from pyopenvba.formula._inventory import excel_has_function
from pyopenvba.formula._parse import FormulaError, Token, literal, parse, split_sheet, tokenize
from pyopenvba.formula._values import number_text

#: The largest number a formula can hold written out.
LARGEST: Final = 9.99999999999999e307
_CORNER: Final = re.compile(r"(\$?)([A-Za-z]{1,3})?(\$?)([0-9]{1,7})?")
#: What Excel reads in a formula that the model does not: structured references and the dynamic-array operators.
_UNMODELLED: Final = re.compile(r"[\[\]@]|#(?!N/A|NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|SPILL!|CALC!|GETTING_DATA)",
                                re.IGNORECASE)
#: A string, or a sheet's name in apostrophes, where any character may stand.
_QUOTED: Final = re.compile(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'')


class UnmodelledFormulaError(FormulaError):
    """A formula Excel reads that uses syntax the model does not."""


class Names(Protocol):
    """What spelling a formula asks of the workbook it goes into."""

    #: The sheet the formula is on, whose names it finds first.
    home: str

    def sheet(self, name: str) -> str | None:
        """A sheet's own name, found in any case, or None."""
        ...

    def defined(self, name: str, sheet: str) -> str | None:
        """A defined name's own spelling as a formula on ``sheet`` would find it, or None."""
        ...

    def remembered(self, name: str) -> str:
        """The spelling the workbook first saw an undefined name in, which from now on is ``name``'s if it has none."""
        ...


def spelled(formula: str, names: Names) -> str:
    """``formula``, which starts with =, as Excel spells it back."""
    body = formula[1:]
    unmodelled = _UNMODELLED.search(_QUOTED.sub('""', body))
    if unmodelled is not None:
        raise UnmodelledFormulaError(f"{unmodelled.group(0)!r} in a formula is not implemented")
    parse(formula)
    tokens = tokenize(body, spaces=True)
    pieces: list[str] = []
    array = 0
    previous: Token | None = None
    index = 0
    while index < len(tokens):
        token = tokens[index]
        after = _next(tokens, index)
        if token.kind == "eof":
            break
        if token.kind == "ws":
            if not array and after.kind not in ("comma", "eof"):
                pieces.append(token.text)
            index += 1
            continue
        if token.kind == "lbrace":
            array += 1
        elif token.kind == "rbrace":
            array -= 1
        text = token.text
        if token.kind == "op" and text in "+-" and _unary(previous) and tokens[index + 1].kind == "number":
            # A sign before a number, with no space between, is the number's own.
            index += 1
            token = tokens[index]
            text = _number(token.text, negative=text == "-")
        elif token.kind == "number":
            text = _number(text, negative=False)
        elif token.kind == "ref":
            text = _reference(text, names)
        elif token.kind == "error":
            head, _, error = text.rpartition("#")
            text = _prefix(head[:-1], names) + "!#" + error.upper() if head else text.upper()
        elif token.kind == "name":
            text = _name(text, tokens[index + 1].kind == "open", names)
        pieces.append(text)
        previous = token
        index += 1
    return "=" + "".join(pieces)


def _next(tokens: list[Token], index: int) -> Token:
    """The first token after ``index`` that is not white space."""
    for token in tokens[index + 1:]:
        if token.kind != "ws":
            return token
    return tokens[-1]


def _unary(previous: Token | None) -> bool:
    """Whether a sign here stands before an operand rather than between two."""
    if previous is None:
        return True
    if previous.kind == "op":
        return previous.text != "%"
    return previous.kind in ("open", "comma", "semicolon", "lbrace")


def _number(text: str, *, negative: bool) -> str:
    value = literal(text)
    if value > LARGEST:
        raise FormulaError(f"{text} is larger than a formula can hold")
    return number_text(-value if negative and value else value, formula=True)


def _prefix(head: str, names: Names) -> str:
    """A sheet written before ``!``, as its own name when the workbook has it."""
    sheet, _ = split_sheet(head + "!")
    found = names.sheet(sheet)
    return head if found is None else quote_sheet(found)


def _reference(text: str, names: Names) -> str:
    head, bang, body = text.rpartition("!")
    corners = [_corner(part) for part in body.split(":")]
    if len(corners) == 2:
        (first_column, first_row), (second_column, second_row) = corners
        first_column, second_column = _ordered(first_column, second_column)
        first_row, second_row = _ordered(first_row, second_row)
        corners = [(first_column, first_row), (second_column, second_row)]
    spelled_body = ":".join(_render(column, row) for column, row in corners)
    return (_prefix(head, names) + "!" if bang else "") + spelled_body


def _corner(text: str) -> tuple[tuple[str, int] | None, tuple[str, int] | None]:
    """One corner of a reference: its column and row, each with its dollar sign, either absent in a whole row or column."""
    found = _CORNER.fullmatch(text)
    assert found is not None
    column_fixed, letters, row_fixed, digits = found.groups()
    column = (column_fixed, column_number(letters)) if letters else None
    # With no letters, as in $1:$2, the only dollar sign is the row's.
    row = (row_fixed or (column_fixed if not letters else ""), int(digits)) if digits else None
    return column, row


def _ordered(first: tuple[str, int] | None, second: tuple[str, int] | None) -> tuple[
        tuple[str, int] | None, tuple[str, int] | None]:
    if first is None or second is None or first[1] <= second[1]:
        return first, second
    return second, first


def _render(column: tuple[str, int] | None, row: tuple[str, int] | None) -> str:
    return (f"{column[0]}{column_letter(column[1])}" if column else "") + (f"{row[0]}{row[1]}" if row else "")


def _name(text: str, called: bool, names: Names) -> str:
    head, bang, bare = text.rpartition("!")
    prefix = _prefix(head, names) + "!" if bang else ""
    if _fixed(bare, called):
        return prefix + bare.upper()
    defined = None if called else names.defined(bare, split_sheet(head + "!")[0] if bang else names.home)
    return prefix + (defined if defined is not None else names.remembered(bare))


def _fixed(bare: str, called: bool) -> bool:
    """Whether a name is one of Excel's own, spelled in capitals: a function it has, or TRUE or FALSE."""
    return excel_has_function(bare) if called else bare.upper() in ("TRUE", "FALSE")


def remembered_names(formula: str, names: Names) -> list[str]:
    """The names in a formula whose spelling a workbook remembers: those it neither defines nor has as functions.

    A workbook read from a file learns them from its formulas, in the order it reads them.
    """
    try:
        tokens = tokenize(formula[1:] if formula.startswith("=") else formula, spaces=True)
    except FormulaError:
        return []
    out: list[str] = []
    for index, token in enumerate(tokens):
        if token.kind != "name":
            continue
        head, bang, bare = token.text.rpartition("!")
        called = tokens[index + 1].kind == "open"
        if _fixed(bare, called):
            continue
        if not called and names.defined(bare, split_sheet(head + "!")[0] if bang else names.home) is not None:
            continue
        out.append(bare)
    return out
