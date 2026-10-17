"""GETPIVOTDATA selects cached displayed report cells, including subtotals.

Adapted from William's pyOfficeEditor, with native duplicate-selector controls.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Protocol, runtime_checkable

from pyopenvba.formula._calc.cells import CellError, UnsupportedFormulaError
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.registry import R, V, function
from pyopenvba.formula._calc.values import REF, Empty, Reference, Scalar, Value

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivot_reports import PivotReport


@runtime_checkable
class PivotSource(Protocol):
    pivot_evaluate: bool

    def pivot_reports(self, sheet: str) -> list[PivotReport]: ...


def _report(context: Context, reference: Reference) -> tuple[str, PivotReport] | None:
    if len(reference.areas) != 1 or not isinstance(context.book, PivotSource):
        return None
    if reference.indirect and context.book.pivot_evaluate:
        return None
    area = reference.areas[0]
    reports = context.book.pivot_reports(area.sheet)
    for report in reports:
        extent = report.extent
        if extent.top <= area.bottom and area.top <= extent.bottom and extent.left <= area.right and area.left <= extent.right:
            return area.sheet, report
    return None


@function('GETPIVOTDATA', V, R, V, V, minimum=2, maximum=254, repeat=2, reject_references=(0,))
def GETPIVOTDATA(context: Context, data_field: Scalar, pivot_table: Value, *pairs: Scalar) -> Value:
    if isinstance(data_field, CellError):
        return data_field
    if isinstance(pivot_table, CellError):
        return pivot_table
    if not isinstance(pivot_table, Reference):
        return REF
    found = _report(context, pivot_table)
    if found is None or len(pairs) % 2:
        return REF
    sheet, report = found
    if not report.readable:
        raise UnsupportedFormulaError('GETPIVOTDATA on a cube or unsupported report layout')
    value = report.value_index(context.text(data_field))
    if value is None:
        return REF
    chosen: dict[int, int] = {}
    for position in range(0, len(pairs), 2):
        field = report.field_index(context.text(pairs[position]))
        wanted = pairs[position + 1]
        if field is None or isinstance(wanted, (Empty, CellError)):
            return REF
        item = report.item_index(field, wanted)
        if item is None or (field in chosen and chosen[field] != item):
            return REF
        chosen[field] = item
    cell = report.cell(value, chosen)
    if cell is None:
        return REF
    shown = context.book.cell(sheet, cell.row, cell.column)
    return 0.0 if isinstance(shown, Empty) else shown
