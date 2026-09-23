"""The functions a formula can call, and how each takes its arguments.

Excel hands an argument over in one of a few ways, and which one a
parameter uses decides what a range given to it does:

- ``VALUE``: one scalar. A range is cut down to the cell in the formula's
  own row or column, the implicit intersection, unless the formula is an
  array formula; an array makes the call run once per item and give an
  array back, which is how ``ABS({-1,-2})`` is ``{1,2}``.
- ``RANGE``: the reference, array or scalar as it is, for a function that
  reads many values, such as SUM, or needs the shape, such as INDEX.
- ``REFERENCE``: a reference and nothing else, for ROW, OFFSET and the like.
- ``LAZY``: the argument's tree, not yet evaluated, for IF and the others
  that evaluate only what they need.

An implementation takes the evaluation context first, then its arguments;
an argument not given is left to the implementation's own default, and one
given empty, as in ``IF(A1,,2)``, is :data:`~.values.EMPTY`. It raises
:class:`~.values.ExcelError` to return an error.

A legacy formula -- one Range.Formula writes, or a file keeps without
``t="array"`` -- is worked out as Excel worked formulas out before
dynamic arrays, and a function declares where that differs, as
pyOpenVBA measured through Range.Formula (tests/fixtures/formula/):
``legacy_first`` are the places that take only the first item of an array
given there, even inside SUMPRODUCT, ``SUM(INDEX({1,2;3,4},{1,2},{1,2}))``
being 1; ``legacy_cell`` the places worked out as a cell works a value
out rather than whole, ``SUM(TRANSPOSE(A1:A3*2))`` being #VALUE! in row 5;
``legacy_corner`` the places that take the first cell of a range, not
the one in the formula's row; and ``legacy_as_cell`` a function that
works all its arguments out as a cell does, even inside an argument
worked out whole, as IF does inside SUMPRODUCT.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from enum import Enum
from typing import TYPE_CHECKING, Any

from pyopenvba.formula._calc.nodes import Missing, Node
from pyopenvba.formula._calc.values import EMPTY, VALUE, Array, ExcelError, Reference, Scalar, Value

if TYPE_CHECKING:
    from pyopenvba.formula._calc.evaluator import Context


class Kind(Enum):
    VALUE = "value"
    RANGE = "range"
    ARRAY = "array"
    REFERENCE = "reference"
    LAZY = "lazy"


V = Kind.VALUE
R = Kind.RANGE
A = Kind.ARRAY
REF = Kind.REFERENCE
LAZY = Kind.LAZY

#: The most arguments Excel lets a call have.
MAX_ARGUMENTS = 255

Implementation = Callable[..., Any]


@dataclass(frozen=True)
class Function:
    name: str
    implementation: Implementation
    kinds: tuple[Kind, ...]
    minimum: int
    maximum: int
    #: How many of the last kinds repeat for arguments past the declared
    #: ones: 1 for SUM's numbers, 2 for SUMIFS's range and criterion pairs.
    repeat: int = 1
    #: Whether a result can change with nothing in the workbook changing,
    #: as NOW's does.
    volatile: bool = False
    #: What a legacy formula does differently, as the module docstring has it.
    legacy_first: frozenset[int] = frozenset()
    legacy_cell: frozenset[int] = frozenset()
    legacy_corner: frozenset[int] = frozenset()
    legacy_as_cell: bool = False

    def kind(self, index: int) -> Kind:
        if index < len(self.kinds):
            return self.kinds[index]
        tail = self.kinds[-self.repeat :]
        return tail[(index - len(self.kinds)) % self.repeat]


FUNCTIONS: dict[str, Function] = {}


def function(
    name: str,
    *kinds: Kind,
    minimum: int | None = None,
    maximum: int | None = None,
    repeat: int = 1,
    volatile: bool = False,
    legacy_first: tuple[int, ...] = (),
    legacy_cell: tuple[int, ...] = (),
    legacy_corner: tuple[int, ...] = (),
    legacy_as_cell: bool = False,
) -> Callable[[Implementation], Implementation]:
    """Register an implementation under ``name``. With no ``minimum`` every
    declared argument is required; with no ``maximum`` no more are allowed.
    """

    def register(implementation: Implementation) -> Implementation:
        FUNCTIONS[name] = Function(
            name,
            implementation,
            kinds,
            len(kinds) if minimum is None else minimum,
            len(kinds) if maximum is None else maximum,
            repeat,
            volatile,
            frozenset(legacy_first),
            frozenset(legacy_cell),
            frozenset(legacy_corner),
            legacy_as_cell,
        )
        return implementation

    return register


def call(entry: Function, context: Context, nodes: tuple[Node, ...]) -> Value:
    """Evaluate the arguments the way each parameter takes them, and call."""
    if context.legacy and context.array and entry.legacy_as_cell:
        # Inside an argument worked out whole, a legacy formula's IF works its own arguments out as a cell does.
        context.array = False
        try:
            return call(entry, context, nodes)
        finally:
            context.array = True
    args: list[Any] = []
    lifted: list[int] = []
    for index, node in enumerate(nodes):
        kind = entry.kind(index)
        if kind is Kind.LAZY:
            args.append(node)
            continue
        if isinstance(node, Missing):
            args.append(EMPTY)
            continue
        whole = kind is Kind.ARRAY and not (context.legacy and index in entry.legacy_cell)
        value = context.evaluate_array(node) if whole else context.evaluate(node)
        if kind is Kind.VALUE:
            if context.legacy and index in entry.legacy_corner and isinstance(value, Reference) \
                    and value.area is not None and not context.array:
                value = context.book.cell(value.area.sheet, value.area.top, value.area.left)
            value = context.operand(value)
            if isinstance(value, Array):
                if context.legacy and index in entry.legacy_first:
                    value = context.first(value)
                else:
                    lifted.append(index)
        elif kind is Kind.REFERENCE and not isinstance(value, Reference):
            return VALUE
        args.append(value)
    if lifted:
        return _lift(entry, context, args, lifted)
    return invoke(entry, context, args)


def invoke(entry: Function, context: Context, args: list[Any]) -> Value:
    try:
        result: Value = entry.implementation(context, *args)
    except ExcelError as error:
        return error.error
    return result


def _lift(entry: Function, context: Context, args: list[Any], positions: list[int]) -> Array:
    """Call once per item of the arrays given for single values."""
    arrays: list[Array] = [args[index] for index in positions]
    height = max(array.height for array in arrays)
    width = max(array.width for array in arrays)
    rows: list[list[Scalar]] = []
    for row in range(height):
        items: list[Scalar] = []
        for column in range(width):
            each = list(args)
            for index in positions:
                each[index] = args[index].at(row, column)
            items.append(context.first(invoke(entry, context, each)))
        rows.append(items)
    return Array(rows)


__all__ = ["FUNCTIONS", "LAZY", "MAX_ARGUMENTS", "REF", "A", "Function", "Kind", "R", "V", "call", "function", "invoke"]
