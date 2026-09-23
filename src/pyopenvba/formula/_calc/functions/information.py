"""The IS functions, TYPE, ERROR.TYPE, NA, and the functions about where a
reference is: ROW, COLUMN, ROWS, COLUMNS, AREAS, ISREF, ISFORMULA."""

from __future__ import annotations

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.common import matrix
from pyopenvba.formula._calc.nodes import Node
from pyopenvba.formula._calc.registry import LAZY, REF, A, R, V, function
from pyopenvba.formula._calc.values import (
    ERROR_NUMBERS,
    NA,
    VALUE,
    Array,
    Empty,
    Reference,
    Scalar,
    Value,
)
from pyopenvba.formula._calc.host import quote_sheet_name
from pyopenvba.formula._calc.reference import CellRef
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.cells import UnsupportedFormulaError


@function("ISBLANK", V)
def ISBLANK(context: Context, value: Scalar) -> Value:
    return isinstance(value, Empty)


@function("ISNUMBER", V)
def ISNUMBER(context: Context, value: Scalar) -> Value:
    return isinstance(value, float)


@function("ISTEXT", V)
def ISTEXT(context: Context, value: Scalar) -> Value:
    return isinstance(value, str)


@function("ISNONTEXT", V)
def ISNONTEXT(context: Context, value: Scalar) -> Value:
    return not isinstance(value, str)


@function("ISLOGICAL", V)
def ISLOGICAL(context: Context, value: Scalar) -> Value:
    return isinstance(value, bool)


@function("ISERROR", V)
def ISERROR(context: Context, value: Scalar) -> Value:
    return isinstance(value, CellError)


@function("ISERR", V)
def ISERR(context: Context, value: Scalar) -> Value:
    return isinstance(value, CellError) and value.code != NA.code


@function("ISNA", V)
def ISNA(context: Context, value: Scalar) -> Value:
    return isinstance(value, CellError) and value.code == NA.code


@function("ISEVEN", V)
def ISEVEN(context: Context, value: Scalar) -> Value:
    return int(context.number(value)) % 2 == 0


@function("ISODD", V)
def ISODD(context: Context, value: Scalar) -> Value:
    return int(context.number(value)) % 2 == 1


@function("ISREF", LAZY)
def ISREF(context: Context, value: Node) -> Value:
    return isinstance(context.evaluate(value), Reference)


@function("ISFORMULA", REF)
def ISFORMULA(context: Context, reference: Reference) -> Value:
    area = reference.areas[0]
    return context.book.formula_text(area.sheet, area.top, area.left) is not None


@function("FORMULATEXT", REF)
def FORMULATEXT(context: Context, reference: Reference) -> Value:
    area = reference.areas[0]
    text = context.book.formula_text(area.sheet, area.top, area.left)
    return NA if text is None else "=" + text


@function("TYPE", LAZY)
def TYPE(context: Context, value: Node) -> Value:
    found = context.evaluate(value)
    if isinstance(found, Reference):
        area = found.area
        if area is None or not area.is_cell:
            return 16.0
        found = context.book.cell(area.sheet, area.top, area.left)
    if isinstance(found, Array):
        return 64.0
    if isinstance(found, bool):
        return 4.0
    if isinstance(found, str):
        return 2.0
    if isinstance(found, CellError):
        return 16.0
    return 1.0


@function("NA")
def NA_(context: Context) -> Value:
    return NA


@function("ERROR.TYPE", V)
def ERROR_TYPE(context: Context, value: Scalar) -> Value:
    if isinstance(value, CellError) and value.code in ERROR_NUMBERS:
        return float(ERROR_NUMBERS[value.code])
    return NA


def _where(context: Context, reference: Value | None, *, row: bool) -> Value:
    if reference is None:
        return float(context.row if row else context.column)
    if not isinstance(reference, Reference) or reference.area is None:
        return VALUE
    area = reference.area
    first = area.top if row else area.left
    count = area.height if row else area.width
    if context.array and count > 1:
        if row:
            return Array([[float(first + index)] for index in range(count)])
        return Array([[float(first + index) for index in range(count)]])
    return float(first)


@function("ROW", R, minimum=0)
def ROW(context: Context, reference: Value | None = None) -> Value:
    return _where(context, reference, row=True)


@function("COLUMN", R, minimum=0)
def COLUMN(context: Context, reference: Value | None = None) -> Value:
    return _where(context, reference, row=False)


@function("ROWS", A)
def ROWS(context: Context, value: Value) -> Value:
    if isinstance(value, Reference):
        if value.area is None:
            return CellError("#REF!")
        return float(value.area.height)
    return float(matrix(context, value).height)


@function("COLUMNS", A)
def COLUMNS(context: Context, value: Value) -> Value:
    if isinstance(value, Reference):
        if value.area is None:
            return CellError("#REF!")
        return float(value.area.width)
    return float(matrix(context, value).width)


@function("AREAS", REF)
def AREAS(context: Context, reference: Reference) -> Value:
    return float(len(reference.areas))


@function("SHEET", R, minimum=0)
def SHEET(context: Context, value: Value | None = None) -> Value:
    order = context.book.sheet_order()
    if value is None:
        return float(order.index(context.sheet) + 1)
    if isinstance(value, Reference):
        return float(order.index(value.areas[0].sheet) + 1)
    if isinstance(value, str):
        found = context.book.sheet_key(value)
        return NA if found is None else float(order.index(found) + 1)
    return VALUE


@function("SHEETS", R, minimum=0)
def SHEETS(context: Context, value: Value | None = None) -> Value:
    if value is None:
        return float(len(context.book.sheet_order()))
    if isinstance(value, Reference):
        return float(len({area.sheet for area in value.areas}))
    return VALUE


#: What CELL tells about a cell's formatting, which this engine does not
#: read: those keep the value Excel cached.
_FORMATTING = frozenset({"color", "format", "parentheses", "prefix", "protect", "width"})


@function("CELL", V, REF, minimum=1)
def CELL(context: Context, info: Scalar, reference: Reference | None = None) -> Value:
    """What CELL tells about a cell: its address, row, column, contents,
    type or file. An address on another sheet names the workbook too, as
    ``[Book1.xlsx]Data!$B$5``, by the name the workbook has now."""
    kind = context.text(info).lower()
    if reference is None:
        raise UnsupportedFormulaError("CELL without a reference, which reads the cell last changed")
    if kind in _FORMATTING:
        raise UnsupportedFormulaError(f'CELL("{kind}"), which reads formatting')
    area = reference.areas[0]
    book = context.book
    if kind == "row":
        return float(area.top)
    if kind == "col":
        return float(area.left)
    if kind == "contents":
        return book.cell(area.sheet, area.top, area.left)
    if kind == "type":
        value = book.cell(area.sheet, area.top, area.left)
        return "b" if isinstance(value, Empty) else "l" if isinstance(value, str) else "v"
    if kind == "address":
        cell = CellRef(area.top, area.left, absolute_row=True, absolute_column=True).a1
        if area.sheet == context.sheet:
            return cell
        # Quoted when the sheet's own name needs it, not for the brackets.
        qualified = f"[{book.name}]{area.sheet}"
        if quote_sheet_name(area.sheet) != area.sheet:
            qualified = "'" + qualified.replace("'", "''") + "'"
        return qualified + "!" + cell
    if kind == "filename":
        if book.folder is None:
            return ""
        return f"{book.folder}[{book.name}]{area.sheet}"
    return VALUE


@function("HYPERLINK", V, V, minimum=1)
def HYPERLINK(context: Context, link: Scalar, name: Scalar | None = None) -> Value:
    """The text shown for a link: its friendly name, or the link itself."""
    if name is None or isinstance(name, Empty):
        return context.text(link)
    if isinstance(name, CellError):
        return name
    return name


__all__: list[str] = []
