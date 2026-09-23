"""The tree a formula parses into.

Every node is immutable, and :func:`render` turns a tree back into formula
text that parses to the same tree. A reference keeps the absolute markers it
was written with, because a formula that is copied or shared moves its
relative parts and not the others.

A few shapes are worth knowing:

- ``A1:B2`` is one :class:`AreaReference`, as it is one token to Excel. The
  range *operator* only appears between things that are not both plain
  cells: ``A1:B2:C3``, ``INDEX(...):D4``, ``Name1:Name2``.
- A sheet qualifier belongs to the reference: ``Data!A1:B2`` is an area on
  ``Data`` with both ends there, and ``Jan:Mar!A1`` is a 3D reference.
- Parentheses are kept, as :class:`Paren`. They do not change a value, but
  Excel keeps them in the formula and one of its rules, the one that turns a
  result a rounding error away from zero into zero, looks at the operation
  a formula ends with.
- An argument left empty, as in ``IF(A1,,2)``, is :class:`Missing`, which is
  not the same as an argument not given at all.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba.formula._calc.host import quote_sheet_name
from pyopenvba.formula._calc.reference import AxisRef, CellRef
from pyopenvba.formula._calc.cells import CellError

if TYPE_CHECKING:
    from pyopenvba.formula._calc.values import Value


@dataclass(frozen=True, slots=True)
class Prefix:
    """What stands before the ``!`` of a reference.

    ``Data!`` has a sheet; ``Jan:Mar!`` a first and last sheet; ``[1]Data!``
    a workbook too, by the index Excel gives each external link; ``[1]!``
    only the workbook, for a name defined there.
    """

    sheet: str | None = None
    last_sheet: str | None = None
    book: str | None = None


@dataclass(frozen=True, slots=True)
class Number:
    value: float
    #: The digits as written, which Excel keeps: ``1E+20`` is stored as
    #: ``100000000000000000000``.
    text: str


@dataclass(frozen=True, slots=True)
class Text:
    value: str


@dataclass(frozen=True, slots=True)
class Logical:
    value: bool


@dataclass(frozen=True, slots=True)
class ErrorLiteral:
    code: str


#: What an array constant may hold: no references and no expressions.
ArrayItem = float | str | bool | CellError


@dataclass(frozen=True, slots=True)
class ArrayLiteral:
    """``{1,2;3,4}``: rows of constants, every row as long as the first."""

    rows: tuple[tuple[ArrayItem, ...], ...]


@dataclass(frozen=True, slots=True)
class CellReference:
    ref: CellRef
    prefix: Prefix | None = None


@dataclass(frozen=True, slots=True)
class AreaReference:
    """A block written as two cells, ``A1:B2``, in the order written."""

    first: CellRef
    last: CellRef
    prefix: Prefix | None = None


@dataclass(frozen=True, slots=True)
class AxisReference:
    """Whole rows or columns: ``2:4`` or ``A:C``."""

    ref: AxisRef
    prefix: Prefix | None = None


@dataclass(frozen=True, slots=True)
class ErrorReference:
    """A reference to cells that were deleted, ``Data!#REF!``."""

    prefix: Prefix | None = None


@dataclass(frozen=True, slots=True)
class NameReference:
    """A defined name, or a table's name on its own."""

    name: str
    prefix: Prefix | None = None


#: The special items of a structured reference, as Excel spells them.
ALL = "#All"
DATA = "#Data"
HEADERS = "#Headers"
TOTALS = "#Totals"
THIS_ROW = "#This Row"
SPECIAL_ITEMS = (ALL, DATA, HEADERS, TOTALS, THIS_ROW)


@dataclass(frozen=True, slots=True)
class StructuredReference:
    """``Table1[[#Headers],[Qty]:[Price]]`` and its shorter forms.

    ``items`` are the special items named, in Excel's spelling; none named
    means the data rows. ``first`` and ``last`` are the columns, the same
    column twice for one, and ``None`` for every column.
    """

    table: str | None
    items: tuple[str, ...] = ()
    first: str | None = None
    last: str | None = None


@dataclass(frozen=True, slots=True)
class Unary:
    """A prefix operator: ``-``, ``+`` or ``@``."""

    op: str
    operand: Node


@dataclass(frozen=True, slots=True)
class Postfix:
    """A postfix operator: ``%``, or ``#`` for a spilled range."""

    op: str
    operand: Node


@dataclass(frozen=True, slots=True)
class Binary:
    """An infix operator.

    Besides the arithmetic, comparison and ``&``, the reference operators:
    ``:`` for the range two references span, ``" "`` for the cells two
    have in common and ``","`` for both together.
    """

    op: str
    left: Node
    right: Node


@dataclass(frozen=True, slots=True)
class Call:
    """A function call, with its name as written: ``_xlfn.XLOOKUP``."""

    name: str
    args: tuple[Node, ...]

    @property
    def function(self) -> str:
        """The name without the prefixes a file adds, in upper case."""
        return function_key(self.name)


@dataclass(frozen=True, slots=True)
class Invoke:
    """A call of what an expression gives, a LAMBDA: ``LAMBDA(x,x*2)(4)``."""

    target: Node
    args: tuple[Node, ...]


@dataclass(frozen=True, slots=True)
class Missing:
    """An argument left empty."""


@dataclass(frozen=True, slots=True, eq=False)
class Given:
    """A value handed to a function from outside any formula, as VBA hands WorksheetFunction one: a scalar, an array
    or cells. pyOpenVBA's own (docs/formula_engine.md)."""

    value: Value


@dataclass(frozen=True, slots=True)
class Paren:
    inner: Node


Node = (
    Number
    | Text
    | Logical
    | ErrorLiteral
    | ArrayLiteral
    | CellReference
    | AreaReference
    | AxisReference
    | ErrorReference
    | NameReference
    | StructuredReference
    | Unary
    | Postfix
    | Binary
    | Call
    | Invoke
    | Missing
    | Paren
    | Given
)

#: The prefixes a file puts before a function newer than Excel 2007, before
#: one of the dynamic-array functions, and before a user-defined function.
_FUNCTION_PREFIXES = ("_xlfn.", "_xlws.", "_xludf.")


def walk(node: Node) -> Iterator[Node]:
    """Every node of a tree, the root first."""
    stack = [node]
    while stack:
        current = stack.pop()
        yield current
        if isinstance(current, (Unary, Postfix)):
            stack.append(current.operand)
        elif isinstance(current, Binary):
            stack.extend((current.right, current.left))
        elif isinstance(current, Call):
            stack.extend(reversed(current.args))
        elif isinstance(current, Invoke):
            stack.extend(reversed(current.args))
            stack.append(current.target)
        elif isinstance(current, Paren):
            stack.append(current.inner)


def function_key(name: str) -> str:
    """A function's name as the registry knows it: ``_xlfn._xlws.SORT`` is
    ``SORT``, and case never matters."""
    key = name.upper()
    stripped = True
    while stripped:
        stripped = False
        for prefix in _FUNCTION_PREFIXES:
            if key.startswith(prefix.upper()):
                key = key[len(prefix) :]
                stripped = True
    return key


# ----------------------------------------------------------------------
# Rendering
# ----------------------------------------------------------------------

#: How tightly each infix operator binds, for deciding where rendering
#: needs parentheses the tree does not record.
BINDING = {
    "=": 10, "<>": 10, "<": 10, ">": 10, "<=": 10, ">=": 10,
    "&": 20,
    "+": 30, "-": 30,
    "*": 40, "/": 40,
    "^": 50,
    ",": 80,
    " ": 90,
    ":": 100,
}  # fmt: skip
PERCENT_BINDING = 60
PREFIX_BINDING = 70
SPILL_BINDING = 110


def render(node: Node) -> str:
    """Formula text for a tree, without the leading ``=``."""
    return _render(node, 0)


def _render(node: Node, context: int) -> str:
    if isinstance(node, Given):
        raise ValueError("a value handed in from outside a formula has no spelling")
    if isinstance(node, Number):
        return node.text
    if isinstance(node, Text):
        return '"' + node.value.replace('"', '""') + '"'
    if isinstance(node, Logical):
        return "TRUE" if node.value else "FALSE"
    if isinstance(node, ErrorLiteral):
        return node.code
    if isinstance(node, ArrayLiteral):
        return "{" + ";".join(",".join(_render_item(item) for item in row) for row in node.rows) + "}"
    if isinstance(node, CellReference):
        return _render_prefix(node.prefix) + node.ref.a1
    if isinstance(node, AreaReference):
        return _render_prefix(node.prefix) + node.first.a1 + ":" + node.last.a1
    if isinstance(node, AxisReference):
        return _render_prefix(node.prefix) + node.ref.a1
    if isinstance(node, ErrorReference):
        return _render_prefix(node.prefix) + "#REF!"
    if isinstance(node, NameReference):
        return _render_prefix(node.prefix) + node.name
    if isinstance(node, StructuredReference):
        return _render_structured(node)
    if isinstance(node, Unary):
        binding = BINDING[":"] - 5 if node.op == "@" else PREFIX_BINDING
        text = node.op + _render(node.operand, binding)
        return f"({text})" if context > binding else text
    if isinstance(node, Postfix):
        binding = PERCENT_BINDING if node.op == "%" else SPILL_BINDING
        text = _render(node.operand, binding) + node.op
        return f"({text})" if context > binding else text
    if isinstance(node, Binary):
        binding = BINDING[node.op]
        # Every infix operator is left-associative, "^" included.
        text = _render(node.left, binding) + node.op + _render(node.right, binding + 1)
        return f"({text})" if context > binding else text
    if isinstance(node, Call):
        return node.name + "(" + ",".join(_render(arg, 0) for arg in node.args) + ")"
    if isinstance(node, Invoke):
        return _render(node.target, SPILL_BINDING) + "(" + ",".join(_render(arg, 0) for arg in node.args) + ")"
    if isinstance(node, Missing):
        return ""
    return "(" + _render(node.inner, 0) + ")"


def _render_item(item: ArrayItem) -> str:
    if isinstance(item, bool):
        return "TRUE" if item else "FALSE"
    if isinstance(item, CellError):
        return item.code
    if isinstance(item, str):
        return '"' + item.replace('"', '""') + '"'
    number = float(item)
    return str(int(number)) if number.is_integer() else repr(number)


def _render_prefix(prefix: Prefix | None) -> str:
    if prefix is None:
        return ""
    book = f"[{prefix.book}]" if prefix.book is not None else ""
    if prefix.sheet is None:
        return book + "!"
    sheets = [prefix.sheet] if prefix.last_sheet is None else [prefix.sheet, prefix.last_sheet]
    text = book + ":".join(sheets)
    if any(quote_sheet_name(sheet) != sheet for sheet in sheets):
        # A 3D pair is quoted as one: 'Q1 Data:Q4 Data'!A1.
        text = "'" + text.replace("'", "''") + "'"
    return text + "!"


def _render_structured(node: StructuredReference) -> str:
    table = node.table or ""
    columns = ""
    if node.first is not None:
        columns = "[" + _escape_column(node.first) + "]"
        if node.last is not None and node.last != node.first:
            columns += ":[" + _escape_column(node.last) + "]"
    if not node.items:
        if not columns:
            return table + "[]"
        return table + columns if node.last in (None, node.first) else table + "[" + columns + "]"
    if len(node.items) == 1 and not columns:
        return table + "[" + node.items[0] + "]"
    parts = ["[" + item + "]" for item in node.items]
    if columns:
        parts.append(columns)
    return table + "[" + ",".join(parts) + "]"


def _escape_column(name: str) -> str:
    """A column name inside brackets: ``'`` escapes the characters that
    would otherwise end or open something."""
    return "".join("'" + char if char in "[]#'" else char for char in name)


__all__ = [
    "ALL",
    "BINDING",
    "DATA",
    "Given",
    "HEADERS",
    "PERCENT_BINDING",
    "PREFIX_BINDING",
    "SPECIAL_ITEMS",
    "SPILL_BINDING",
    "THIS_ROW",
    "TOTALS",
    "AreaReference",
    "ArrayItem",
    "ArrayLiteral",
    "AxisReference",
    "Binary",
    "Call",
    "CellReference",
    "ErrorLiteral",
    "ErrorReference",
    "Invoke",
    "Logical",
    "Missing",
    "NameReference",
    "Node",
    "Number",
    "Paren",
    "Postfix",
    "Prefix",
    "StructuredReference",
    "Text",
    "Unary",
    "function_key",
    "render",
    "walk",
]
