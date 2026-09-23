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
- A structured reference is spelled as :mod:`pyopenvba.formula._structured`
  sets out, naming its table even inside it, and where one value is
  wanted a column of a table with more than one row of data becomes this
  row's cell of it (scripts/measure_structured_references.py).

What Excel refuses to read raises :class:`FormulaError`; what it reads
and the model does not -- the ``@`` and ``#`` of dynamic arrays, a
reference to another workbook -- raises :class:`UnmodelledFormulaError`.
"""

from __future__ import annotations

import re
from dataclasses import replace
from typing import Final, Protocol

from pyopenvba._a1 import column_letter, column_number, quote_sheet
from pyopenvba.formula._functions import AS_CELL, CELLS, FUNCTIONS, REFERENCES, array_places, one_value_places
from pyopenvba.formula._inventory import excel_has_function
from pyopenvba.formula._parse import (REFERENCE_OPS, Binary, Call, FormulaError, NameNode, Node, Reference, Structured,
                                      Token, Unary, literal, parse, read_structured, split_sheet, tokenize)
from pyopenvba.formula._structured import TableShape, one_cell, spelled as spelled_reference
from pyopenvba.formula._values import number_text

#: The largest number a formula can hold written out.
LARGEST: Final = 9.99999999999999e307
_CORNER: Final = re.compile(r"(\$?)([A-Za-z]{1,3})?(\$?)([0-9]{1,7})?")
#: What Excel reads in a formula that the model does not: the dynamic-array operators.
_UNMODELLED: Final = re.compile(r"@|#(?!N/A|NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|SPILL!|CALC!|GETTING_DATA)",
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

    def table(self, name: str) -> TableShape | None:
        """The table of that name, found in any case, or None."""
        ...

    def here(self) -> TableShape | None:
        """The table the formula's cell is in, which a reference leaving out its table's name means, or None."""
        ...


def spelled(formula: str, names: Names, *, whole: bool = False) -> str:
    """``formula``, which starts with =, as Excel spells it back; ``whole`` for a formula worked out whole, as an
    array formula is, where no column is cut to this row."""
    body = formula[1:]
    try:
        tokens = tokenize(body, spaces=True)
    except FormulaError:
        unmodelled = _UNMODELLED.search(_QUOTED.sub('""', body))
        if unmodelled is not None:
            raise UnmodelledFormulaError(f"{unmodelled.group(0)!r} in a formula is not implemented") from None
        raise
    tree = parse(formula)
    _check(tree)
    one_value: set[int] = set()
    if not whole:
        _one_value(tree, one=True, whole=False, found=one_value)
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
        elif token.kind == "structured":
            following = tokens[index + 1]
            if following.kind in ("ref", "name", "structured") and following.at == token.at + len(token.text):
                # [Book2]Sheet1!A1: another workbook's cells.
                raise UnmodelledFormulaError("a reference to another workbook is not implemented")
            text = _structured(token, token.at in one_value, names)
        elif token.kind == "error":
            head, _, error = text.rpartition("#")
            text = _prefix(head[:-1], names) + "!#" + error.upper() if head else text.upper()
        elif token.kind == "name":
            text = _name(text, tokens[index + 1].kind == "open", names)
        pieces.append(text)
        previous = token
        index += 1
    return "=" + "".join(pieces)


#: The fewest arguments a function that reads cells is taken with. Measured (scripts/measure_subtotal.py):
#: SUBTOTAL(9) is error 1004 when written.
_FEWEST: Final = {"SUBTOTAL": 2}


def _check(node: Node | None) -> None:
    """Refuse what Excel refuses in a formula past its syntax: a value where a function reads cells.

    Which arguments have to be cells is :data:`~pyopenvba.formula._functions.CELLS`, measured in
    tests/fixtures/formula/probes.txt and by scripts/measure_subtotal.py.
    """
    if isinstance(node, Call):
        name = node.name.upper()
        places = CELLS.get(name)
        if places is not None and (len(node.args) < _FEWEST.get(name, 0)
                                   or not all(_referring(node.args[index]) for index in places.within(len(node.args)))):
            raise FormulaError(f"{node.name} takes references")
        for argument in node.args:
            _check(argument)
    elif isinstance(node, Unary):
        _check(node.operand)
    elif isinstance(node, Binary):
        _check(node.left)
        _check(node.right)


def _referring(node: Node | None) -> bool:
    """Whether an argument can stand for cells: a reference, a name, references joined, or a call.

    A call counts only to a function that can answer with cells: OFFSET,
    INDEX or IF can, LEN, SUM or IFERROR cannot. A function the model does
    not have is let through, since what it answers with is not known here.
    """
    if isinstance(node, (Reference, Structured)):
        return True
    if isinstance(node, Call):
        name = node.name.upper()
        return name in REFERENCES or name not in FUNCTIONS
    if isinstance(node, NameNode):
        return node.name.upper() not in ("TRUE", "FALSE")
    if isinstance(node, Binary):
        return node.op in (":", " ", ",") and _referring(node.left) and _referring(node.right)
    return False


def _one_value(node: Node | None, *, one: bool, whole: bool, found: set[int]) -> None:
    """Where the structured references stand that a cell cuts to its own row: each place intersected works out
    as one value (see pyopenvba.formula._engine), outside an argument worked out whole."""
    if isinstance(node, Structured):
        if one and not whole:
            found.add(node.at)
    elif isinstance(node, Unary):
        _one_value(node.operand, one=True, whole=whole, found=found)
    elif isinstance(node, Binary):
        cells = node.op in REFERENCE_OPS
        _one_value(node.left, one=not cells, whole=whole, found=found)
        _one_value(node.right, one=not cells, whole=whole, found=found)
    elif isinstance(node, Call):
        count = len(node.args)
        places, arrays = one_value_places(node.name, count), array_places(node.name, count)
        inside = False if node.name.upper() in AS_CELL else whole
        for index, argument in enumerate(node.args):
            _one_value(argument, one=index in places, whole=inside or index in arrays, found=found)


def _structured(token: Token, one: bool, names: Names) -> str:
    """A structured reference as Excel spells it back, naming its table; Excel refuses one to a table or a column
    the workbook has not got."""
    node = read_structured(token.text)
    table = names.table(node.table) if node.table else names.here()
    if table is None:
        raise FormulaError(f"no table for {token.text!r}")
    if node.first is not None:
        first, last = table.column(node.first), table.column(node.last or node.first)
        if first is None or last is None:
            raise FormulaError(f"{table.name} has no column for {token.text!r}")
        node = replace(node, first=table.columns[first], last=table.columns[last])
    if one and not node.items and node.first is not None and not node.span and table.data_rows > 1:
        node = one_cell(node)
    return spelled_reference(replace(node, table=table.name, sheet=""))


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
