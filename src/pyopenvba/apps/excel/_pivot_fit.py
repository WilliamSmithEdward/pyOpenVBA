"""How Excel fits a pivot table's columns each time it lays the table out, Autofit column widths on update.

Measured with scripts/measure_pivot_column_fit.py (tests/fixtures/pivot_column_fit/), and checked against every
pivot sheet the earlier fixtures saved. A table whose useAutoFormatting holds, HasAutoFormat, makes each of its
columns as wide as AutoFit makes a column for the cells of the table in it, narrower as well as wider, a width a
macro gave the column too; nothing outside the table counts, and a column the table left keeps its width. An
empty table fits nothing. The column of a row field that is neither compact nor the innermost row field keeps its
width as well. Where the columns show their fields' captions, each run of item columns that a subtotal's or the
grand total's column closes is as wide as the widest of them; a run the table's edge ends is not. A change of
style fits the columns again; the style's flags and the +/- buttons change nothing until the table is next laid
out.

A cell is measured as AutoFit measures it (see _column_fit), in the bold of every element of the table's style
that reaches it and draws bold, as Excel's TableStyleElements(...).Font.Bold reads them for each built-in style:
the header rows, the grand total row and column, the first column, the first header cell, each depth's subtotal
rows and columns and row and column subheadings, and the page fields' labels and values; an inner field's total
over every outer item takes no subtotal's. The style's flags turn some off: ShowTableStyleColumnHeaders the
header rows and the column subheadings, ShowTableStyleRowHeaders the first column and the row subheadings,
ShowTableStyleLastColumn the grand total column.

A cell showing a field's dropdown is 16 pixels wider, after the width its text makes: Row Labels, Column Labels,
a field's caption over its column and a page field's value, not the data field's caption in the corner, the Values
field's nor a page field's own caption. An item with a field inside it, not the Values field alone, shows a +/-
button ahead of its label: a column item's label starts 88/7 pixels in; in a compact label column holding such
items every label does, and each depth a further 9 pixels for each step of CompactRowIndent, 90/7 at least, where
no step leaves it where its parent starts; with no buttons a depth is 9 pixels a step, as an indent is. A
subtotal's or the grand total's label starts where its indent puts it, the buttons or none. The two constants are
fitted, not derived: they hold for every one of 1,280 measured labels.
"""
from __future__ import annotations

from collections.abc import Sequence
from copy import deepcopy
from dataclasses import dataclass
from fractions import Fraction
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._pivots import NS
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._pivots import PivotTable

# The table style elements a pivot table's cells take, as XlTableStyleElementType numbers them.
HEADER_ROW, GRAND_TOTAL_ROW, FIRST_COLUMN, GRAND_TOTAL_COLUMN = 1, 2, 3, 4
FIRST_HEADER_CELL = 9
SUBTOTAL_COLUMN, SUBTOTAL_ROW = 13, 16
COLUMN_SUBHEADING, ROW_SUBHEADING = 20, 23
PAGE_LABELS, PAGE_VALUES = 26, 27
#: The pixels a field's dropdown adds to a cell, after the width its text makes.
DROPDOWN = 16
#: Where a compact label starts beside the +/- buttons, and the least a depth adds there (see the module's notes).
BUTTON, LEAST_STEP = Fraction(88, 7), Fraction(90, 7)
#: The pixels an indent's step is.
INDENT = 9

def _numbered(prefix: str, kind: str, numbers: str) -> list[str]:
    """Style names: ``prefix`` and ``kind`` with each number, or range of them, in ``numbers``."""
    names: list[str] = []
    for part in numbers.split():
        first, _, last = part.partition('-')
        names.extend(f'{prefix}{kind}{number}' for number in range(int(first), int(last or first) + 1))
    return names


#: The elements each built-in pivot table and table style draws bold, by the styles that share them
#: (tests/fixtures/pivot_column_fit/ styles).
_BOLD_STYLES: dict[tuple[int, ...], list[str]] = {
    (1, 2, 16, 17, 23, 24): _numbered('PivotStyle', 'Light', '1 3-21'),
    (1, 2, 16, 23): _numbered('PivotStyle', 'Light', '2'),
    (1, 2, 3, 13, 16, 17, 26): _numbered('PivotStyle', 'Light', '22'),
    (): _numbered('PivotStyle', 'Light', '23-28'),
    (2, 9, 16, 17): _numbered('PivotStyle', 'Medium', '1-7'),
    (1, 2, 13, 16, 17, 23, 24): _numbered('PivotStyle', 'Medium', '8-14'),
    (16, 17, 23, 24, 25): _numbered('PivotStyle', 'Medium', '15-21'),
    (1, 2, 3, 16, 17, 23, 26): _numbered('PivotStyle', 'Medium', '22-28'),
    (1, 2, 16, 23, 24, 26): _numbered('PivotStyle', 'Dark', '1-7 15-21'),
    (1, 2, 16, 23, 24): _numbered('PivotStyle', 'Dark', '8-14'),
    (2, 9, 16, 17, 23, 24): _numbered('PivotStyle', 'Dark', '22'),
    (2, 9, 17, 23, 24): _numbered('PivotStyle', 'Dark', '23-28'),
    (1, 2, 3, 4): [*_numbered('TableStyle', 'Light', '1-21'), *_numbered('TableStyle', 'Medium', '1-14 22-28'),
                   *_numbered('TableStyle', 'Dark', '1-7')],
    (1, 3, 4): _numbered('TableStyle', 'Medium', '15-21'),
    (2, 3, 4): _numbered('TableStyle', 'Dark', '8-11'),
}
BOLD: dict[str, frozenset[int]] = {name.casefold(): frozenset(elements)
                                   for elements, names in _BOLD_STYLES.items() for name in names}


@dataclass(frozen=True)
class Fit:
    """What a table cell asks of its column: the style elements that reach it, how far into the cell its text
    starts, in pixels, and the pixels its dropdown adds."""

    elements: frozenset[int] = frozenset()
    offset: Fraction = Fraction(0)
    extra: int = 0


def label_offset(depth: int, indent: int, buttons: bool, *, total: bool = False) -> Fraction:
    """How far into its cell a compact row label at ``depth`` starts (see the module's notes); a subtotal's or the
    grand total's label starts where its cell's indent puts it, the buttons or none."""
    if total or not buttons:
        return Fraction(INDENT * indent * depth)
    return BUTTON + depth * (max(Fraction(INDENT * indent), LEAST_STEP) if indent else Fraction(0))


def _shown_elements(root: XmlElement) -> frozenset[int]:
    """The bold elements the table's style draws, less those its flags turn off."""
    info = root.find(NS + 'pivotTableStyleInfo')
    if info is None:
        return frozenset()
    bold = set(BOLD.get(info.get('name', '').casefold(), frozenset()))

    def shown(name: str, default: bool) -> bool:
        return info.get(name, '1' if default else '0') in ('1', 'true')

    if not shown('showColHeaders', False):
        bold -= {HEADER_ROW, COLUMN_SUBHEADING, COLUMN_SUBHEADING + 1, COLUMN_SUBHEADING + 2}
    if not shown('showRowHeaders', False):
        bold -= {FIRST_COLUMN, ROW_SUBHEADING, ROW_SUBHEADING + 1, ROW_SUBHEADING + 2}
    if not shown('showLastColumn', False):
        bold.discard(GRAND_TOTAL_COLUMN)
    return frozenset(bold)


def refit(table: PivotTable) -> None:
    """Fit a table's columns again for a change that lays out no cell, a style or its flags."""
    from pyopenvba.apps.excel._pivot_layout import plan_render

    try:
        plan = plan_render(table, deepcopy(table.root))
    except (VBARuntimeError, VBAUnsupportedError):
        return
    fit_columns(table.sheet, table.root, plan.fits, plan.fitted, plan.groups)


def fit_columns(sheet: Worksheet, root: XmlElement, fits: dict[tuple[int, int], Fit], columns: set[int],
                groups: Sequence[Sequence[int]] = ()) -> None:
    """Fit ``columns`` of a table laid out on ``sheet`` to its cells, each measured as ``fits`` has it, each group
    of ``groups`` as wide as the widest of its columns."""
    from pyopenvba.apps.excel import _column_fit
    from pyopenvba.apps.excel._font_widths import widths

    if root.get('useAutoFormatting', '0') not in ('1', 'true') or not sheet.dims.normal_known():
        return
    bold = _shown_elements(root)
    normal = sheet.book.stylesheet.default.font
    widest: dict[int, int] = {}
    refused: set[int] = set()
    for (row, column), fit in fits.items():
        if column not in columns or column in refused:
            continue
        cell = sheet.cells_.get((row, column))
        if cell is None:
            continue
        try:
            found = _column_fit.shown_value(sheet, row, column, cell)
        except VBAUnsupportedError:
            refused.add(column)
            continue
        if found is None:
            continue
        shown = found[1]
        font = sheet.style_at(row, column).font
        table = widths(font.name, font.bold or bool(fit.elements & bold), font.italic, float(font.size))
        if (table is None or font.name != normal.name or font.size != normal.size or font.vert_align
                or any(char not in table for char in shown.text + shown.pads) or '\n' in shown.text):
            # Measured only in the Normal font AutoFit measures; a column holding anything else keeps its width.
            refused.add(column)
            continue
        pixels = _column_fit.column_pixels(Fraction(_column_fit.measured_width(table, shown)) + fit.offset,
                                           _column_fit.font_pad(table)) + fit.extra
        widest[column] = max(widest.get(column, 0), pixels)
    for group in groups:
        measured = [widest[column] for column in group if column in widest and column not in refused]
        for column in group:
            if measured and column in columns and column not in refused:
                widest[column] = max(measured)
    for column, pixels in widest.items():
        if column not in refused:
            sheet.dims.fit_column(column, min(pixels, _column_fit.WIDEST))
    if widest:
        sheet.touched()
