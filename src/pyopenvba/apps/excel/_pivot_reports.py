"""Saved PivotTable report selection, adapted from William's pyOfficeEditor.

Axis mappings select actual displayed report cells, rather than recomputing
totals from source records. This retains filtered/hidden/subtotal behavior.
"""
from __future__ import annotations

import datetime as dt
import math
from dataclasses import dataclass
from collections.abc import Callable, Iterator
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._pivots import decode_text as decode
from pyopenvba.formula._calc import collate as _collate
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.dates import serial
from pyopenvba.formula._calc.host import format_value
from pyopenvba.formula._calc.reference import CellRef, RangeRef

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotTable


def datetime_to_serial(value: dt.datetime, *, epoch_1904: bool = False) -> float:
    return float(serial(value.year, value.month, value.day, epoch_1904=epoch_1904)) + (
        value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400


class Element:
    """Read-only navigation over the hardened reader's already parsed nodes."""
    def __init__(self, node: XmlElement) -> None:
        self.node = node

    @property
    def name(self) -> str:
        return self.node.tag.rpartition('}')[2]

    def get(self, name: str) -> str | None:
        return self.node.get(name)

    def elements(self) -> Iterator[Element]:
        return (Element(node) for node in self.node)

    def children_named(self, name: str) -> Iterator[Element]:
        return (entry for entry in self.elements() if entry.name == name)

    def child(self, name: str) -> Element | None:
        return next(self.children_named(name), None)


VALUES = -2

#: What turns on a field's subtotals besides its default one.
_SUBTOTAL_FLAGS = (
    "sumSubtotal", "countASubtotal", "avgSubtotal", "maxSubtotal", "minSubtotal", "productSubtotal",
    "countSubtotal", "stdDevSubtotal", "stdDevPSubtotal", "varSubtotal", "varPSubtotal",
)  # fmt: skip

Item = str | float | bool | CellError | None


@dataclass(frozen=True)
class PivotItem:
    """An item of a field: its value in the cache, a date as its serial
    number, and where the cache lists it; the text the report shows it as;
    and the name it was given in the report, if it was."""

    value: Item
    index: int
    text: str
    caption: str | None


@dataclass(frozen=True)
class PivotField:
    """A field of a report: the name the report gives it, its items in
    the report's order, ``None`` for an entry that is a subtotal, and
    whether it shows subtotals and whether at the top of each group.

    A grouped field says how: ``years``, ``quarters`` or ``months`` of a
    date, or ``range`` for numbers in bins of ``interval`` from ``start``.
    """

    name: str
    items: tuple[PivotItem | None, ...]
    subtotals: bool
    subtotals_top: bool
    grouping: str | None = None
    start: float = 0.0
    interval: float = 0.0

    def number_index(self, wanted: float) -> int | None:
        """The item a number picks. Measured: a date grouped by years,
        quarters or months takes the whole number part as the year, the
        quarter or the month; a bin is found by whole steps from its start,
        toward zero, and takes a number only at or below the bin's start;
        anything else has to be equal."""
        target: int | None = None
        if self.grouping == "range":
            if self.interval <= 0:
                return None
            steps = math.trunc((wanted - self.start) / self.interval)
            if wanted > self.start + steps * self.interval:
                return None
            target = steps + 1
        elif self.grouping in ("years", "quarters", "months"):
            target = math.trunc(wanted)
        for index, item in enumerate(self.items):
            if item is None:
                continue
            if target is None:
                if isinstance(item.value, float) and item.value == wanted:
                    return index
            elif self.grouping == "years":
                if item.text.isdigit() and int(item.text) == target:
                    return index
            elif item.index == target:
                return index
        return None


#: The item a line stands at in a field it skips, one inside a collapsed item (see _lines).
GAP = -1
#: The item a line stands at in the fields outside an inner field's subtotal over every item of theirs.
EVERY = 1048832


@dataclass(frozen=True)
class PivotLine:
    """A row or column of a report's data area: its kind (``data``,
    ``grand``, ``blank`` or a subtotal's function), the item each field of
    its axis stands at, the outer fields first, and the value it shows
    when its axis lists the values."""

    kind: str
    items: tuple[int, ...]
    data: int


@dataclass(frozen=True)
class PivotReport:
    """What a pivot table shows and where: the fields, the lines of its
    data area as the file lists them, its filters, and its values, each
    by its name and the name of the field it summarizes."""

    name: str
    location: RangeRef
    #: The location with the filters above it, which name the report too.
    extent: RangeRef
    first_data_row: int
    first_data_column: int
    fields: tuple[PivotField, ...]
    rows: tuple[int, ...]
    columns: tuple[int, ...]
    row_lines: tuple[PivotLine, ...]
    column_lines: tuple[PivotLine, ...]
    #: Each filter's field, with the item it is set to, or ``None`` when it
    #: shows every item.
    pages: tuple[tuple[int, int | None], ...]
    values: tuple[tuple[str, str], ...]
    #: Whether the report is one this reads: not one from a cube, whose
    #: fields GETPIVOTDATA names another way, and with its lines written.
    readable: bool = True

    def value_index(self, name: str) -> int | None:
        """The value a name picks: its own name, or else the field it
        summarizes, the first such."""
        for index, (own, _) in enumerate(self.values):
            if _collate.equal(own, name):
                return index
        for index, (_, source) in enumerate(self.values):
            if _collate.equal(source, name):
                return index
        return None

    def field_index(self, name: str) -> int | None:
        for index, field in enumerate(self.fields):
            if _collate.equal(field.name, name):
                return index
        return None

    def item_index(self, field: int, wanted: str | float | bool) -> int | None:
        """The item a value picks: text by the name it was given or the
        text it is shown as, a number by its value."""
        found = self.fields[field]
        if isinstance(wanted, float):
            return found.number_index(wanted)
        for index, item in enumerate(found.items):
            if item is None:
                continue
            if isinstance(wanted, str):
                if (item.caption is not None and _collate.equal(item.caption, wanted)) or _collate.equal(item.text, wanted):
                    return index
            elif isinstance(item.value, bool) and item.value == wanted:
                return index
        return None

    def cell(self, value: int, chosen: dict[int, int]) -> CellRef | None:
        """The cell of the data area showing a value for the items chosen,
        a field left out meaning its total, or ``None`` where the report
        shows no such cell. A filter chosen has to be set to that item."""
        remaining = dict(chosen)
        for field, item in self.pages:
            if field in remaining and (item is None or remaining.pop(field) != item):
                return None
        if any(field not in self.rows and field not in self.columns for field in remaining):
            return None
        row = self._line(self.row_lines, self.rows, remaining, value)
        column = self._line(self.column_lines, self.columns, remaining, value)
        if row is None or column is None:
            return None
        top = self.location.top + self.first_data_row
        left = self.location.left + self.first_data_column
        return CellRef(top + row, left + column)

    def _line(self, lines: tuple[PivotLine, ...], axis: tuple[int, ...], chosen: dict[int, int], value: int) -> int | None:
        """The line of an axis that stands at the items chosen of its
        fields and shows the value wanted, when the axis lists the values.

        Measured: the line with no other field at an item, the total of the
        fields left out, is the one; failing that, a line whose other
        fields are at items, if it is the only one. A group's own line,
        which leaves the fields within it out, shows its subtotal only when
        the field has subtotals at the top."""
        values_at = axis.index(VALUES) if VALUES in axis else None
        wanted = {position: chosen[field] for position, field in enumerate(axis) if field in chosen}
        others: list[int] = []
        for index, line in enumerate(lines):
            if line.kind == "blank":
                continue
            assigned = 0 if line.kind == "grand" else len(line.items)
            if values_at is not None:
                if values_at < assigned:
                    shown = line.items[values_at]
                elif line.kind == "data":
                    continue
                else:
                    shown = line.data
                if shown != value:
                    continue
            if any(position >= assigned or line.items[position] != item for position, item in wanted.items()):
                continue
            if line.kind == "data" and assigned < len(axis):
                inner = axis[assigned - 1] if assigned else VALUES
                if inner == VALUES or not (self.fields[inner].subtotals and self.fields[inner].subtotals_top):
                    continue
            if all(position in wanted or position == values_at for position in range(assigned)):
                return index
            others.append(index)
        return others[0] if len(others) == 1 else None


def read_report(table: PivotTable, formats: Callable[[int], str], *, epoch_1904: bool = False) -> PivotReport | None:
    """Read the saved report geometry and cached axis items without recalculating it."""
    root = Element(table.root)
    location = root.child("location")
    if location is None:
        return None
    try:
        block = RangeRef.parse(location.get("ref") or "").normalized
    except ValueError:
        return None
    pages_rows = _number(location.get("rowPageCount"))
    top = max(block.top - pages_rows - 1, 1) if pages_rows else block.top
    extent = RangeRef(CellRef(top, block.left), CellRef(block.bottom, block.right))
    unread = PivotReport(decode(root.get("name") or ""), block, extent, 0, 0, (), (), (), (), (), (), (), readable=False)
    definition = Element(table.cache.root)
    container = definition.child("cacheFields")
    sources = [] if container is None else list(container.children_named("cacheField"))
    elements = _children(root, "pivotFields", "pivotField")
    rows = tuple(_number(entry.get("x")) for entry in _children(root, "rowFields", "field"))
    columns = tuple(_number(entry.get("x")) for entry in _children(root, "colFields", "field"))
    if (
        definition.child("cacheHierarchies") is not None
        or len(elements) != len(sources)
        or (rows and root.child("rowItems") is None)
        or (columns and root.child("colItems") is None)
    ):
        return unread
    fields = tuple(
        _field(element, source, formats, epoch_1904=epoch_1904) for element, source in zip(elements, sources, strict=True)
    )
    values = tuple(
        (decode(entry.get("name") or ""), decode(sources[_number(entry.get("fld"))].get("name") or ""))
        for entry in _children(root, "dataFields", "dataField")
        if 0 <= _number(entry.get("fld")) < len(sources)
    )
    return PivotReport(
        name=unread.name,
        location=block,
        extent=extent,
        first_data_row=_number(location.get("firstDataRow")),
        first_data_column=_number(location.get("firstDataCol")),
        fields=fields,
        rows=rows,
        columns=columns,
        row_lines=_lines(root.child("rowItems")),
        column_lines=_lines(root.child("colItems")),
        pages=tuple(
            (_number(entry.get("fld")), None if entry.get("item") is None else _number(entry.get("item")))
            for entry in _children(root, "pageFields", "pageField")
        ),
        values=values,
    )


def _children(root: Element, container: str, name: str) -> list[Element]:
    found = root.child(container)
    return [] if found is None else list(found.children_named(name))


def _number(raw: str | None) -> int:
    try:
        return int(raw or 0)
    except ValueError:
        return 0


def axis_lines(root: XmlElement, axis: str) -> tuple[PivotLine, ...]:
    """A table definition's lines of ``rowItems`` or ``colItems``, as the file lists them, whatever else the table
    holds."""
    found = Element(root).child(axis)
    return _lines(found)


def _lines(container: Element | None) -> tuple[PivotLine, ...]:
    """An axis's lines, each ``<i>`` repeating the first ``r`` items of the
    one before it; a grand total's placeholder item left off. A collapsed
    item has no lines for the fields inside it, so the line after its own
    can repeat places its line never had: those stand at GAP
    (tests/fixtures/pivot_cells/ pivot_show_detail.json values_rows)."""
    if container is None:
        return ()
    found: list[PivotLine] = []
    previous: list[int] = []
    for line in container.children_named("i"):
        kind = line.get("t") or "data"
        repeated = _number(line.get("r"))
        kept = previous[:repeated]
        items = kept + [GAP] * (repeated - len(kept)) + [_number(entry.get("v")) for entry in line.children_named("x")]
        found.append(PivotLine(kind, () if kind == "grand" else tuple(items), _number(line.get("i"))))
        previous = items
    return tuple(found)


def _field(element: Element, source: Element, formats: Callable[[int], str], *, epoch_1904: bool) -> PivotField:
    """A field as the report has it, its items read from the cache: from
    the groups a grouped field makes, or else from the values it holds."""
    group = source.child("fieldGroup")
    listed = None if group is None else group.child("groupItems")
    ranged = None if listed is None or group is None else group.child("rangePr")
    if listed is None:
        listed = source.child("sharedItems")
    cached = [] if listed is None else [_cached(entry, epoch_1904=epoch_1904) for entry in listed.elements()]
    code = formats(_number(element.get("numFmtId") or source.get("numFmtId")))
    items: list[PivotItem | None] = []
    for entry in _children(element, "items", "item"):
        index = _number(entry.get("x"))
        if (entry.get("t") or "data") != "data" or not 0 <= index < len(cached):
            items.append(None)
            continue
        value = cached[index]
        caption = entry.get("n")
        text = _shown(value, code, epoch_1904=epoch_1904)
        items.append(PivotItem(value, index, text, None if caption is None else decode(caption)))
    subtotals = element.get("defaultSubtotal") != "0" or any(element.get(flag) in ("1", "true") for flag in _SUBTOTAL_FLAGS)
    grouping = None if ranged is None else ranged.get("groupBy") or "range"
    return PivotField(
        name=decode(element.get("name") or source.get("name") or ""),
        items=tuple(items),
        subtotals=subtotals,
        subtotals_top=element.get("subtotalTop") not in ("0", "false"),
        grouping=grouping,
        start=0.0 if ranged is None else _decimal(ranged.get("startNum")),
        interval=0.0 if ranged is None else _decimal(ranged.get("groupInterval")),
    )


def _decimal(raw: str | None) -> float:
    try:
        return float(raw or 0)
    except ValueError:
        return 0.0


def _cached(entry: Element, *, epoch_1904: bool) -> Item:
    """A cached value: text, a number, a logical, a date as its serial, an
    error, or ``None`` for a blank."""
    raw = decode(entry.get("v") or "")
    kind = entry.name
    if kind == "n":
        try:
            return float(raw)
        except ValueError:
            return None
    if kind == "b":
        return raw in ("1", "true")
    if kind == "d":
        try:
            return datetime_to_serial(dt.datetime.fromisoformat(raw), epoch_1904=epoch_1904)
        except ValueError:
            return None
    if kind == "e":
        return CellError(raw)
    if kind == "s":
        return raw
    return None


def _shown(value: Item, code: str, *, epoch_1904: bool) -> str:
    """The text a report shows an item as."""
    if value is None:
        return "(blank)"
    if isinstance(value, CellError):
        return value.code
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return format_value(float(value), code, epoch_1904=epoch_1904)
    return value
