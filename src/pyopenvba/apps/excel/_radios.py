"""Excel Forms radio groups, in stored control order."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, VBAErrorValue
from pyopenvba.shapes import Shape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet


def groups(sheet: Worksheet) -> list[list[Shape]]:
    result: dict[int, list[Shape]] = {}
    for item in sheet.shapes_:
        if item.control is not None and item.control.kind == "Radio":
            if item.control.radio_group is None:
                item.control.radio_group = _container(sheet, item)
            result.setdefault(item.control.radio_group, []).append(item)
    return list(result.values())


def _contains(box: Shape, item: Shape) -> bool:
    return (box.left <= item.left and box.top <= item.top
            and item.left + item.width <= box.left + box.width
            and item.top + item.height <= box.top + box.height)


def _container(sheet: Worksheet, item: Shape, excluding: Shape | None = None) -> int:
    matches = [box for box in sheet.shapes_ if box is not excluding
               and box.control and box.control.kind == "GBox" and _contains(box, item)]
    if not matches:
        return 0
    winner = matches[0]
    for box in matches[1:]:
        # A strictly nested box wins. Partial/identical overlaps retain the
        # earlier box, even when a later partial overlap has a smaller area.
        if _contains(winner, box) and not _contains(box, winner):
            winner = box
    return winner.shape_id


def initialize(sheet: Worksheet) -> None:
    """Resolve membership when opening, before any geometry edits."""
    groups(sheet)


def group(sheet: Worksheet, shape: Shape) -> list[Shape]:
    return next(items for items in groups(sheet) if any(one is shape for one in items))


def link(sheet: Worksheet, shape: Shape) -> str:
    control = group(sheet, shape)[0].control
    assert control is not None
    return control.linked_cell


def selected(items: list[Shape]) -> int:
    return next((i for i, item in enumerate(items, 1) if item.control and item.control.value == 1), 0)


def select(sheet: Worksheet, items: list[Shape], index: int) -> None:
    for i, item in enumerate(items, 1):
        assert item.control is not None
        wanted = 1 if i == index else -4146
        if item.control.value != wanted:
            item.control.value = wanted
            sheet.drawing_changed()


def cell_value(sheet: Worksheet, items: list[Shape], value: object) -> None:
    if value is EMPTY or isinstance(value, (ExcelError, VBAErrorValue)):
        select(sheet, items, 0)
    elif isinstance(value, (bool, int, float)):
        select(sheet, items, int(value))


def write_link(sheet: Worksheet, reference: str, value: int, *, only_if_changed: bool = False) -> None:
    from pyopenvba.apps.excel._controls import linked_target

    target = linked_target(sheet, reference)
    if target:
        other, row, column = target
        old = other.book.calculator.value_of(other.name, row, column)
        if only_if_changed and (0 if old is EMPTY else old) == value:
            return
        cell = other.cell(row, column, create=True)
        assert cell is not None
        cell.value, cell.formula, cell.stale = float(value), "", False
        other.cell_changed(row, column)


def set_value(sheet: Worksheet, shape: Shape, value: int) -> None:
    from pyopenvba.apps.excel._controls import linked_target, refresh

    if value not in (-4146, 0, 1):
        raise ValueError("radio value must be -4146 or 0 (off), or 1 (on)")
    items = group(sheet, shape)
    reference = link(sheet, shape)
    linked_target(sheet, reference)
    refresh(sheet, shape)
    assert shape.control is not None
    wanted = 1 if value == 1 else -4146
    if shape.control.value == wanted:
        return
    index = next(i for i, item in enumerate(items, 1) if item is shape) if wanted == 1 else 0
    select(sheet, items, index)
    write_link(sheet, reference, index)


def set_link(sheet: Worksheet, shape: Shape, reference: str) -> None:
    from pyopenvba.apps.excel._controls import linked_target, refresh

    linked_target(sheet, reference)
    refresh(sheet, shape)
    items = group(sheet, shape)
    control = items[0].control
    assert control is not None
    control.linked_cell = reference
    sheet.drawing_changed()
    write_link(sheet, reference, selected(items), only_if_changed=True)


def creation_group(sheet: Worksheet, kind: int, left: float, top: float, width: float, height: float) -> bool:
    """Validate grouping before adding a control; return firstButton."""
    existing = groups(sheet)
    candidate = Shape(left=left, top=top, width=width, height=height)
    if kind == 4:
        from pyopenvba.apps.excel._controls import linked_target

        if not any(_contains(candidate, item) for items in existing for item in items):
            return False
        for items in existing:
            control = items[0].control
            assert control is not None
            linked_target(sheet, control.linked_cell)
        return False
    if kind != 7:
        return False
    wanted = _container(sheet, candidate)
    return not any(items[0].control and items[0].control.radio_group == wanted for items in existing)


def _regroup(sheet: Worksheet, excluding: Shape | None = None, *, adding: Shape | None = None,
             sync_links: bool = True) -> None:
    """Rebuild groups in drawing order after a box change."""
    from pyopenvba.apps.excel._controls import linked_target

    previous_links: dict[int, str] = {}
    previous_sizes: dict[int, int] = {}
    previous_groups: dict[int, int] = {}
    for items in groups(sheet):
        control = items[0].control
        assert control is not None and control.radio_group is not None
        previous_links[control.radio_group] = control.linked_cell
        previous_sizes[control.radio_group] = len(items)
        for item in items:
            previous_groups[item.shape_id] = control.radio_group
    radios = [item for item in sheet.shapes_ if item.control and item.control.kind == "Radio"]
    destinations = [_container(sheet, item, excluding) for item in radios]
    transferred_to_sheet = False
    if adding is not None and 0 in previous_links and not previous_links[0]:
        identical = next((box for box in sheet.shapes_ if box is not adding and box.control
                          and box.control.kind == "GBox" and _contains(box, adding) and _contains(adding, box)), None)
        if identical is not None and previous_links.get(identical.shape_id):
            # Excel gives an unlinked sheet group the duplicate box's link.
            previous_links[0] = previous_links[identical.shape_id]
            previous_links[identical.shape_id] = ""
            transferred_to_sheet = True
    # Validate every reference before mutating the group structure.
    for item in radios:
        assert item.control is not None
        linked_target(sheet, item.control.linked_cell)
    for item, destination in zip(radios, destinations):
        assert item.control is not None
        item.control.radio_group = destination
    for items in groups(sheet):
        index = selected(items)
        leader = items[0].control
        assert leader is not None and leader.radio_group is not None
        if excluding is not None:
            destination = leader.radio_group
            if destination and destination in previous_links:
                leader.linked_cell = previous_links[destination]
            else:
                leader.linked_cell = previous_links[previous_groups[items[0].shape_id]]
        elif adding is not None:
            destination = leader.radio_group
            old_owner = previous_groups[items[0].shape_id]
            if destination in previous_links and (destination or transferred_to_sheet):
                leader.linked_cell = previous_links[destination]
            elif old_owner and old_owner in destinations and destination != old_owner:
                # A nested box captures buttons, but the surviving original
                # group retains its linked cell even if its leader changed.
                leader.linked_cell = ""
        for i, item in enumerate(items):
            assert item.control is not None
            item.control.first_button = i == 0
            if i:
                item.control.linked_cell = ""
        select(sheet, items, index)
    if sync_links:
        for items in groups(sheet):
            control = items[0].control
            assert control is not None
            if (excluding is not None and control.radio_group == 0
                    and len(items) < previous_sizes[previous_groups[items[0].shape_id]]):
                # Removing a partially overlapping box can leave only part
                # of its old group unboxed. Excel leaves that cell untouched.
                continue
            write_link(sheet, control.linked_cell, selected(items), only_if_changed=True)
    sheet.drawing_changed()


def added_box(sheet: Worksheet, box: Shape) -> None:
    inside = [_contains(box, item) for item in sheet.shapes_ if item.control and item.control.kind == "Radio"]
    if any(inside):
        # Excel preserves old cell indexes when the split interleaves groups.
        interleaved = sum(a != b for a, b in zip(inside, inside[1:])) > 1
        _regroup(sheet, adding=box, sync_links=not interleaved)


def deleting(sheet: Worksheet, shape: Shape) -> None:
    """Keep group boundaries on deletion; Excel leaves the cell unchanged."""
    if shape.control is None:
        return
    if shape.control.kind == "GBox":
        if any(_contains(shape, item) for items in groups(sheet) for item in items):
            _regroup(sheet, excluding=shape)
    elif shape.control.kind == "Radio":
        items = group(sheet, shape)
        if items[0] is shape and len(items) > 1:
            successor = items[1].control
            assert successor is not None
            successor.first_button = True
            boxed = bool(shape.control.radio_group)
            successor.linked_cell = shape.control.linked_cell if boxed else ""
