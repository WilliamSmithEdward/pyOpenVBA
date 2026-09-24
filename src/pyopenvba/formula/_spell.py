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
from dataclasses import dataclass, replace
from typing import Final, Protocol

from pyopenvba._a1 import column_letter, column_number, quote_sheet
from pyopenvba.formula._arity import takes
from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc.catalog import is_excel_function
from pyopenvba.formula._calc.nodes import function_key
from pyopenvba.formula._calc.registry import FUNCTIONS
from pyopenvba.formula._deep import deep
from pyopenvba.formula._parse import (REFERENCE_OPS, Binary, Call, FormulaError, Invoke, NameNode, Node, Reference,
                                      Structured, Token, Unary, literal, parse, read_structured, split_sheet, tokenize)
from pyopenvba.formula._prefixes import bound
from pyopenvba.formula._r1c1 import is_cell, refused_in_a1
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


@dataclass(frozen=True)
class _Places:
    """Which of a function's arguments a rule covers: some places, and every ``step``-th from ``start`` on."""

    fixed: tuple[int, ...] = ()
    start: int | None = None
    step: int = 1

    def within(self, count: int) -> frozenset[int]:
        run = range(self.start, count, self.step) if self.start is not None else range(0)
        return frozenset(index for index in (*self.fixed, *run) if index < count)


_EVERY: Final = _Places(start=0)

# Measured in tests/fixtures/formula/probes.txt, one probe or more per function.

#: Where a function wants one value and runs item by item through an array given there, its answer an array:
#: SUM(LEN({"a","bb"})) is 3. In a cell, cells given there are first cut to the formula's own row or column.
_LIFTED: Final[dict[str, _Places]] = {
    **dict.fromkeys((
        "ABS", "ADDRESS", "CEILING", "CEILING.MATH", "CHAR", "CODE", "CONCATENATE", "DATE", "DATEVALUE", "DAY",
        "DAYS", "EDATE", "EOMONTH", "ERROR.TYPE", "EXACT", "EXP", "FIND", "FLOOR", "FLOOR.MATH", "HOUR", "INT",
        "ISBLANK", "ISERR", "ISERROR", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISTEXT", "LEFT", "LEN", "LN",
        "LOG", "LOG10", "LOWER", "MID", "MINUTE", "MOD", "MONTH", "MROUND", "NOT", "POWER", "PROPER", "RANDBETWEEN",
        "REPLACE", "REPT", "RIGHT", "ROUND", "ROUNDDOWN", "ROUNDUP", "SEARCH", "SECOND", "SIGN", "SQRT", "SUBSTITUTE",
        "TEXT", "TIME", "TRIM", "TRUNC", "UPPER", "VALUE", "WEEKDAY", "YEAR"), _EVERY),
    "COUNTIF": _Places((1,)), "SUMIF": _Places((1,)), "AVERAGEIF": _Places((1,)),
    "COUNTIFS": _Places(start=1, step=2), "SUMIFS": _Places(start=2, step=2),
    "MATCH": _Places((0, 2)), "LARGE": _Places((1,)), "SMALL": _Places((1,)),
}
#: Where a function wants one value but takes only the first item of an array: SUM(INDEX({1,2;3,4},{1,2},{1,2}))
#: is 1, even inside SUMPRODUCT. Cells are cut as for the functions above.
_FIRST_ITEM: Final[dict[str, _Places]] = {
    "INDEX": _Places((1, 2, 3)), "VLOOKUP": _Places((0, 2, 3)), "HLOOKUP": _Places((0, 2, 3)),
    "XLOOKUP": _Places((0,)),
}
#: Where the functions that work out their own arguments read one value through intersected.
_OWN_PLACES: Final[dict[str, _Places]] = {
    "IF": _Places((0,)), "IFS": _Places(start=0, step=2), "IFERROR": _EVERY, "IFNA": _EVERY, "CHOOSE": _Places((0,)),
    "OFFSET": _Places(start=1),
}
#: Arguments worked out as arrays even in a cell: SUMPRODUCT(LEN(A1:A3)) adds every length, and INDEX(A1:A3*2,2)
#: is A2*2 in any row.
_ARRAYS: Final[dict[str, _Places]] = {"SUMPRODUCT": _EVERY, "INDEX": _Places((0,))}
#: Arguments that have to be cells. Excel will not take a formula with a value there (error 1004): a number, text,
#: an array, an operator's answer, or a function that answers with a value, SUMIF(LEN(A1:A3),1). A name or a
#: function that can answer with cells is taken, and is #VALUE! when it comes to a value instead. Each function
#: Excel refused an operation in was written with one in each argument in turn (scripts/measure_formula_refusals.py,
#: tests/fixtures/formula_refusals.json); AGGREGATE wants cells only from its fifth argument, where it can only be
#: summing references, and a database function only its database.
_CELLS: Final[dict[str, _Places]] = {
    "SUBTOTAL": _Places(start=1), "COUNTIF": _Places((0,)), "SUMIF": _Places((0, 2)), "AVERAGEIF": _Places((0, 2)),
    "COUNTIFS": _Places(start=0, step=2), "SUMIFS": _Places((0,), start=1, step=2),
    "AVERAGEIFS": _Places((0,), start=1, step=2), "MAXIFS": _Places((0,), start=1, step=2),
    "MINIFS": _Places((0,), start=1, step=2), "COUNTBLANK": _Places((0,)), "OFFSET": _Places((0,)),
    "ROW": _Places((0,)), "COLUMN": _Places((0,)), "AREAS": _Places((0,)), "AGGREGATE": _Places(start=4),
    "CELL": _Places((1,)), "RANK": _Places((1,)), "RANK.AVG": _Places((1,)), "RANK.EQ": _Places((1,)),
    "FORMULATEXT": _Places((0,)), "ISFORMULA": _Places((0,)), "PHONETIC": _Places((0,)),
    **dict.fromkeys(("DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM",
                     "DVAR", "DVARP"), _Places((0,))),
}
#: The functions that can answer with cells, the only ones Excel takes where cells are wanted: every other function
#: the engine has is refused there (scripts/measure_cells_functions.py, tests/fixtures/formula/cells_functions.json).
_ANSWER_CELLS: Final = frozenset({"CHOOSE", "DROP", "IF", "IFS", "INDEX", "INDIRECT", "LAMBDA", "LET", "OFFSET",
                                  "REDUCE", "SINGLE", "SWITCH", "TAKE", "TRIMRANGE", "XLOOKUP"})
#: The functions that work their arguments out as a cell does even inside an argument worked out as an array.
_AS_CELL: Final = frozenset({"IF", "IFS", "IFERROR", "IFNA", "SWITCH", "CHOOSE"})


def _places(table: dict[str, _Places], name: str, count: int) -> frozenset[int]:
    places = table.get(name)
    return frozenset() if places is None else places.within(count)


def _one_value_places(name: str, count: int) -> frozenset[int]:
    """The arguments of a call that want one value, cells there cut to the formula's own row or column."""
    upper = name.upper()
    if upper == "SWITCH":
        # The subject and each case; a value, and the default after the last case, can be cells.
        return frozenset({0, *range(1, count - 1, 2)})
    return _places(_LIFTED, upper, count) | _places(_FIRST_ITEM, upper, count) | _places(_OWN_PLACES, upper, count)


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


def spelled(formula: str, names: Names, *, whole: bool = False, at: bool = False, r1c1: bool = False) -> str:
    """``formula``, which starts with =, as Excel spells it back; ``whole`` for a formula worked out whole, as an
    array formula is, where no column is cut to this row; ``at`` for a formula written through Formula2, which may
    cut a range to one value with @; ``r1c1`` for one read in R1C1 and turned into A1, where a LET's or a LAMBDA's
    name may be a cell, x1, which A1 cannot read."""
    return deep(lambda: _spelled(formula, names, whole=whole, at=at, r1c1=r1c1))


def _spelled(formula: str, names: Names, *, whole: bool, at: bool, r1c1: bool) -> str:
    body = formula[1:]
    try:
        tokens = tokenize(body, spaces=True)
    except FormulaError:
        unmodelled = _UNMODELLED.search(_QUOTED.sub('""', body))
        if unmodelled is not None:
            raise UnmodelledFormulaError(f"{unmodelled.group(0)!r} in a formula is not implemented") from None
        raise
    if not at and any(token.kind == "op" and token.text == "@" for token in tokens):
        raise UnmodelledFormulaError("'@' in a formula written through Range.Formula is not implemented")
    _within_limits(tokens)
    bound_at = bound(body)
    if not r1c1 and any(token.kind == "name" and token.at in bound_at and is_cell(token.text) for token in tokens):
        # LET(x1,5,x1): A1 cannot read a cell as a name, and Excel reads the formula as R1C1, where x1 is one and so
        # is every other cell written as A1 (tests/fixtures/bound_names.json).
        raise FormulaError("A1 cannot read a cell as a name a LET or a LAMBDA binds")
    tree = parse(formula)
    _check(tree)
    one_value: set[int] = set()
    if not whole:
        _one_value(tree, one=True, whole=False, found=one_value)
    pieces: list[str] = []
    declared: dict[str, str] = {}
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
        elif token.kind == "structured" and token.at in bound_at:
            # A LAMBDA's parameter a call may leave out, [y], as it is written.
            pass
        elif token.kind == "structured":
            following = tokens[index + 1]
            if following.kind in ("ref", "name", "structured") and following.at == token.at + len(token.text):
                # [Book2]Sheet1!A1: another workbook's cells.
                raise UnmodelledFormulaError("a reference to another workbook is not implemented")
            text = _structured(token, token.at in one_value, names)
        elif token.kind == "error":
            head, _, error = text.rpartition("#")
            text = _prefix(head[:-1], names) + "!#" + error.upper() if head else text.upper()
        elif token.kind == "name" and token.at in bound_at:
            # A name a LET or a LAMBDA binds is spelled everywhere as the formula first writes it, where it is bound:
            # LET(a1,5,A1) reads back as LET(a1,5,a1) (tests/fixtures/bound_names.json).
            text = declared.setdefault(text.upper(), text)
        elif token.kind == "name":
            text = _name(text, tokens[index + 1].kind == "open", names)
        pieces.append(text)
        previous = token
        index += 1
    return "=" + "".join(pieces)


def _check(node: Node | None) -> None:
    """Refuse what Excel refuses in a formula past its syntax: a function given too few or too many arguments, a
    value where a function reads cells, and LET's and LAMBDA's names badly named.

    How many arguments each function takes is :func:`pyopenvba.formula._arity.takes`, and which arguments have to
    be cells :data:`_CELLS`, measured in tests/fixtures/formula/probes.txt and
    tests/fixtures/formula_refusals.json.
    """
    if isinstance(node, Call):
        name = function_key(node.name)
        count = len(node.args)
        if name == "ANCHORARRAY":
            # The file's spelling of A1#, which Excel does not take written as a call (tests/fixtures/formula/).
            raise FormulaError("ANCHORARRAY is not taken written as a call")
        if name in ("LET", "LAMBDA"):
            _check_names(node, name)
        elif not takes(name, count):
            raise FormulaError(f"{node.name} does not take {count} arguments")
        places = _CELLS.get(name)
        if places is not None and not all(_referring(node.args[index]) for index in places.within(count)):
            raise FormulaError(f"{node.name} takes references")
        for argument in node.args:
            _check(argument)
    elif isinstance(node, Invoke):
        for argument in [node.target, *node.args]:
            _check(argument)
    elif isinstance(node, Unary):
        _check(node.operand)
    elif isinstance(node, Binary):
        _check(node.left)
        _check(node.right)


def _referring(node: Node | None) -> bool:
    """Whether an argument can stand for cells: a reference, a name, references joined, or a call.

    A call counts only to a function that can answer with cells: OFFSET,
    INDEX or IF can, LEN, SUM or IFERROR cannot. A function the engine
    has not got is let through, since what it answers with is not known
    here.
    """
    if isinstance(node, (Reference, Structured)) or isinstance(node, Unary) and node.op == "#":
        return True
    if isinstance(node, Call):
        name = function_key(node.name)
        return name in _ANSWER_CELLS or name not in FUNCTIONS
    if isinstance(node, NameNode):
        return node.name.upper() not in ("TRUE", "FALSE")
    if isinstance(node, Invoke):
        # A LAMBDA called where it is written, COUNTIF(LAMBDA(x,x)(A1),1), is taken (cells_functions.json).
        return True
    if isinstance(node, Binary):
        return node.op in (":", " ", ",") and _referring(node.left) and _referring(node.right)
    return False


#: The most arguments LET and LAMBDA take: LET's names and values in pairs with its calculation last, LAMBDA's
#: parameters with its body last.
_MOST_BOUND: Final = {"LET": 253, "LAMBDA": 254}


def _check_names(node: Call, name: str) -> None:
    """Refuse a LET or a LAMBDA Excel refuses: LET without a calculation after its pairs, a LAMBDA with nothing,
    or a name that is not one or is there twice."""
    count = len(node.args)
    if name == "LET" and (count < 3 or count % 2 == 0) or name == "LAMBDA" and count < 1 or count > _MOST_BOUND[name]:
        raise FormulaError(f"{name} does not take {count} arguments")
    names = node.args[0:-1:2] if name == "LET" else node.args[:-1]
    seen: set[str] = set()
    for argument in names:
        bound = _bound(argument, optional=name == "LAMBDA")
        if bound is None or bound.upper() in seen:
            raise FormulaError(f"{name} cannot bind that name")
        seen.add(bound.upper())


def _bound(node: Node | None, *, optional: bool) -> str | None:
    """The name a LET or a LAMBDA binds where ``node`` stands, or None for one Excel does not take: a number, TRUE,
    a name with a full stop in it, or cells, $A$1 or A1:A2; ``optional`` for LAMBDA, whose [name] may be left out.
    One cell with no $, x1, is a name the tokenizer has already made one (pyopenvba.formula._parse._bind_cells)."""
    if isinstance(node, NameNode) and not node.sheet:
        return None if node.name.upper() in ("TRUE", "FALSE") or "." in node.name else node.name
    if optional and isinstance(node, Structured) and not node.table and not node.sheet and not node.items \
            and node.first is not None and node.first == node.last and "." not in node.first:
        return node.first
    return None


#: The most characters a text in a formula holds; the most brackets open in one argument, a call's own bracket
#: among them; the most calls one inside another; and the most signs and operators waiting on their operands in
#: one bracket (scripts/measure_formula_refusals.py, tests/fixtures/formula_refusals.json).
_LONGEST_TEXT: Final = 4095
_MOST_BRACKETS: Final = 256
_MOST_CALLS: Final = 65
_MOST_WAITING: Final = 1024
#: How tightly each operator binds, as Excel's operator stack sorts them out.
_BINDING: Final = {"^": 4, "*": 3, "/": 3, "+": 2, "-": 2, "&": 1, "=": 0, "<>": 0, "<": 0, ">": 0, "<=": 0, ">=": 0,
                   ":": 7, ",": 6}
_SIGN: Final = 5


def _within_limits(tokens: list[Token]) -> None:
    """Refuse a formula past Excel's limits: a text too long, brackets or calls too deep, or too many signs and
    operators waiting on their operands.

    Each argument of a call counts its brackets afresh; the call's own
    bracket counts in the argument it is in. Each bracket, a call's too,
    waits on operators afresh.
    """
    brackets = [0]
    waiting: list[list[int]] = [[]]
    #: What each open bracket is: a call's, a bracket around a part of the formula, or an array constant's.
    opened: list[str] = []
    calls = 0
    operand = False
    previous: Token | None = None
    for token in tokens:
        kind = token.kind
        if kind in ("ws", "eof"):
            continue
        if kind == "text" and len(token.text) - 2 - token.text[1:-1].count('""') > _LONGEST_TEXT:
            raise FormulaError("a text in a formula holds 4095 characters")
        if kind == "open":
            call = previous is not None and previous.kind in ("name", "close")
            brackets[-1] += 1
            if brackets[-1] > _MOST_BRACKETS:
                raise FormulaError("brackets go 256 deep")
            if call:
                calls += 1
                if calls > _MOST_CALLS:
                    raise FormulaError("calls go 65 deep")
                brackets.append(0)
            opened.append("call" if call else "bracket")
            waiting.append([])
            operand = False
        elif kind == "lbrace":
            opened.append("array")
            waiting.append([])
            operand = False
        elif kind in ("close", "rbrace"):
            closing = opened.pop() if opened else ""
            if closing == "call":
                brackets.pop()
                calls -= 1
            if closing in ("call", "bracket") and brackets[-1]:
                brackets[-1] -= 1
            if len(waiting) > 1:
                waiting.pop()
            operand = True
        elif kind in ("comma", "semicolon") and opened:
            waiting[-1].clear()
            operand = False
        elif kind == "op":
            stack = waiting[-1]
            if token.text == "#":
                # What spilled from a cell, E1#, is part of its operand.
                pass
            elif token.text == "%":
                while stack and stack[-1] >= _SIGN:
                    stack.pop()
            elif not operand:
                stack.append(_SIGN)
            else:
                binding = _BINDING.get(token.text, 0)
                while stack and stack[-1] >= binding:
                    stack.pop()
                stack.append(binding)
                operand = False
            if len(stack) > _MOST_WAITING:
                raise FormulaError("too many operators wait on their operands")
        else:
            operand = True
        previous = token


def _one_value(node: Node | None, *, one: bool, whole: bool, found: set[int]) -> None:
    """Where the structured references stand that a cell cuts to its own row: each place intersected works out
    as one value (see pyopenvba.formula._calc.evaluator), outside an argument worked out whole."""
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
        places, arrays = _one_value_places(node.name, count), _places(_ARRAYS, node.name.upper(), count)
        inside = False if node.name.upper() in _AS_CELL else whole
        for index, argument in enumerate(node.args):
            _one_value(argument, one=index in places, whole=inside or index in arrays, found=found)
    elif isinstance(node, Invoke):
        for argument in [node.target, *node.args]:
            _one_value(argument, one=False, whole=whole, found=found)


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
        # A percent sign and a spill's # follow an operand, E1#+1.
        return previous.text not in ("%", "#")
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
    """A name as Excel spells it back, one a LET or a LAMBDA binds aside."""
    head, bang, bare = text.rpartition("!")
    prefix = _prefix(head, names) + "!" if bang else ""
    if bare.startswith("'"):
        # A name that looks like a cell, which R1C1 read as a name: in quotes, as the workbook first saw it.
        return prefix + "'" + names.remembered(_unquoted(bare)).replace("'", "''") + "'"
    if not called and refused_in_a1(bare):
        # R1C1, RC or R: A1 cannot read it, and Excel reads the formula as R1C1 instead.
        raise FormulaError(f"A1 cannot read the name {bare!r}")
    if _fixed(bare, called):
        return prefix + bare.upper()
    defined = None if called else names.defined(bare, split_sheet(head + "!")[0] if bang else names.home)
    return prefix + (defined if defined is not None else names.remembered(bare))


def _unquoted(bare: str) -> str:
    """A name in quotes, 'A1', without them."""
    return bare[1:-1].replace("''", "'") if bare.startswith("'") else bare


def _fixed(bare: str, called: bool) -> bool:
    """Whether a name is one of Excel's own, spelled in capitals: a function a sheet has, SIN and NORM.DIST as much
    as SUM, or TRUE or FALSE (tests/fixtures/formula_spelling.json)."""
    return is_excel_function(bare) if called else bare.upper() in ("TRUE", "FALSE")


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
        bare = _unquoted(bare)
        if not called and names.defined(bare, split_sheet(head + "!")[0] if bang else names.home) is not None:
            continue
        out.append(bare)
    return out
