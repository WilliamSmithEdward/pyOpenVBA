"""Measured checkbox/list values and linked-cell synchronization."""
from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Sequence

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, VBAErrorValue
from pyopenvba.shapes import Shape
from pyopenvba.shapes._values import ControlInfo

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet


def linked_target(sheet: Worksheet, reference: str) -> tuple[Worksheet, int, int] | None:
    from pyopenvba.apps.excel._control_refs import binding

    try:
        _, area = binding(sheet, reference)
    except ValueError as exc:
        raise VBAUnsupportedError(str(exc)) from exc
    if area is None:
        return None
    target = next((one for one in sheet.book.sheets_
                   if one.name.lower() == (area.sheet or sheet.name).lower()), None)
    if target is None:
        raise VBAUnsupportedError(f"control linked worksheet {area.sheet!r} is unavailable")
    return target, area.top, area.left


def value_control(shape: Shape) -> ControlInfo:
    control = shape.control
    if control is None or control.kind not in {"CheckBox", "Drop", "List", "Spin", "Scroll", "Radio"}:
        raise VBAUnsupportedError("control values support checkboxes, lists, spinners and scroll bars")
    if control.kind in {"Drop", "List"} and control.selection_mode != "single":
        raise VBAUnsupportedError("multi-selection list values are not implemented")
    return control


def list_count(sheet: Worksheet, shape: Shape) -> int:
    from pyopenvba.apps.excel._control_refs import binding

    control = list_control(shape)
    if control.kind not in {"Drop", "List"}:
        raise VBAUnsupportedError("ListCount requires a dropdown or list box")
    if not control.list_range:
        return len(control.items)
    _, area = binding(sheet, control.list_range)
    return area.rows if area else 0


def list_control(shape: Shape) -> ControlInfo:
    control = shape.control
    if control is None or control.kind not in {"Drop", "List"}:
        raise VBAUnsupportedError("this operation requires a dropdown or list box")
    return control


def items(sheet: Worksheet, shape: Shape) -> list[str]:
    from pyopenvba.apps.excel._control_refs import binding
    from pyopenvba.apps.excel._calc import as_vba
    from pyopenvba.interpreter._values import to_text

    control = list_control(shape)
    if not control.list_range:
        return list(control.items)
    list_count(sheet, shape)  # validate the source
    _, area = binding(sheet, control.list_range)
    if area is None:
        return []
    source = next(one for one in sheet.book.sheets_ if one.name.lower() == (area.sheet or sheet.name).lower())
    return [to_text(as_vba(source.book.calculator.value_of(source.name, row, area.left)))
            for row in range(area.top, area.bottom + 1)]


def edit_items(sheet: Worksheet, shape: Shape, operation: str, index: int = 0,
               text: str = "", count: int = 1) -> None:
    from pyopenvba.apps.excel._shape_api import validate_text

    control = list_control(shape)
    validate_text(text)
    if operation == "set" and index < 1:
        raise ValueError("list item index must be positive")
    if control.list_range:
        if operation == "remove":
            raise ValueError("Excel cannot remove an individual item from a range-backed list")
        # Excel discards the source binding rather than copying its rows.
        # Resetting its selection also writes zero to a single-select link.
        set_value(sheet, shape, 0)
        control.list_range = ""
        control.items.clear()
    size = len(control.items)
    if operation == "remove" and not 1 <= index <= size:
        raise ValueError("list item index is outside the list")
    if operation == "remove" and (count < 1 or index + count - 1 > size):
        raise ValueError("list item removal is outside the list")
    if operation == "add":
        at = index if 1 <= index <= size else size + 1
        control.items.insert(at - 1, text)
        if control.value >= at:
            control.value += 1
        control.selected_indices = [one + 1 if one >= at else one for one in control.selected_indices]
    elif operation == "remove":
        del control.items[index - 1:index - 1 + count]
        def shifted(one: int) -> int:
            return one - count if one >= index + count else (0 if one >= index else one)
        control.value = shifted(control.value)
        control.selected_indices = [shifted(one) for one in control.selected_indices if shifted(one)]
    elif operation == "set":
        if index > size:
            control.items.append(text)
        else:
            control.items[index - 1] = text
    elif operation == "clear":
        control.items.clear()
        control.value = 0
        control.selected_indices.clear()
    else:
        raise ValueError("unknown list edit")
    sheet.drawing_changed()


def replace_items(sheet: Worksheet, shape: Shape, entries: Sequence[object]) -> None:
    """Python whole-list replacement validates all strings before mutation."""
    from pyopenvba.apps.excel._shape_api import validate_text

    prepared: list[str] = []
    for entry in entries:
        if not isinstance(entry, str):
            raise ValueError("list items must be strings")
        validate_text(entry)
        prepared.append(entry)
    edit_items(sheet, shape, "clear")
    for entry in prepared:
        edit_items(sheet, shape, "add", text=entry)


def set_mode(sheet: Worksheet, shape: Shape, mode: int) -> None:
    control = list_control(shape)
    mode = {-4142: 1, -4154: 2}.get(mode, mode)
    if mode not in (1, 2, 3):
        raise ValueError("selection mode must be 1, 2 or 3")
    if control.kind == "Drop":
        return
    wanted = {1: "single", 2: "multi", 3: "extended"}[mode]
    if wanted != control.selection_mode:
        if "single" in {wanted, control.selection_mode}:
            control.selected_indices.clear()
        control.selection_mode = wanted
        if wanted == "single":
            control.value = 0
        sheet.drawing_changed()


def set_selection(sheet: Worksheet, shape: Shape, indices: Sequence[object]) -> None:
    control = list_control(shape)
    if control.selection_mode == "single":
        raise ValueError("multiple selected indexes require multi or extended mode")
    count = list_count(sheet, shape)
    selected: set[int] = set()
    for one in indices:
        if not isinstance(one, int) or not 1 <= one <= count:
            raise ValueError("selection index is outside the list")
        selected.add(one)
    control.selected_indices = sorted(selected)
    sheet.drawing_changed()


def linked_value(control: ControlInfo) -> object:
    if control.kind in {"Drop", "List", "Spin", "Scroll"}:
        return float(control.value)
    return ExcelError("#N/A") if control.value == 2 else control.value == 1


def list_selection(value: object, count: int) -> int | None:
    if value is EMPTY or isinstance(value, (ExcelError, VBAErrorValue)):
        return 0
    if isinstance(value, (bool, int, float)):
        return min(max(int(value), 0), count)
    return None


def list_source_changed(sheet: Worksheet, shape: Shape) -> None:
    """A changed source reinterprets the linked index against the new row count."""
    control = shape.control
    assert control is not None
    if control.selection_mode != "single":
        count = list_count(sheet, shape)
        control.selected_indices = [one for one in control.selected_indices if one <= count]
        return
    count = list_count(sheet, shape)
    control.value = min(control.value, count)
    try:
        target = linked_target(sheet, control.linked_cell)
    except VBAUnsupportedError:
        return
    if target:
        other, row, column = target
        selected = list_selection(other.book.calculator.value_of(other.name, row, column), count)
        if selected is not None:
            control.value = selected


def refresh(sheet: Worksheet, shape: Shape) -> None:
    """Calculate a linked formula before reading or assigning the control."""
    control = shape.control
    if control is None or control.kind not in {"CheckBox", "Drop", "List", "Spin", "Scroll", "Radio"}:
        return
    try:
        if control.kind in {"Drop", "List"} and control.list_range:
            count = list_count(sheet, shape)
            selected = [one for one in control.selected_indices if one <= count]
            value = min(control.value, count)
            if selected != control.selected_indices or value != control.value:
                control.selected_indices, control.value = selected, value
                sheet.drawing_changed()
        reference = control.linked_cell
        if control.kind == "Radio":
            from pyopenvba.apps.excel._radios import link

            reference = link(sheet, shape)
        target = linked_target(sheet, reference)
    except (ValueError, VBAUnsupportedError):
        # Reading or saving an existing unsupported binding preserves it.
        # Explicit assignments still validate the binding before mutation.
        return
    if target:
        other, row, column = target
        # Direct cell writes already notified controls. Calculating a stale
        # formula notifies them when its result becomes available.
        other.book.calculator.value_of(other.name, row, column)


def set_value(sheet: Worksheet, shape: Shape, value: object) -> None:
    control = shape.control
    if control is None or control.kind not in {"Drop", "List"}:
        control = value_control(shape)
    if not isinstance(value, int):
        raise ValueError("control value must be an integer")
    if control.kind == "Radio":
        from pyopenvba.apps.excel._radios import set_value as set_radio_value

        set_radio_value(sheet, shape, value)
        return
    if control.kind == "CheckBox":
        if value not in (-4146, 0, 1, 2):
            raise ValueError("checkbox value must be -4146 (off), 0 (off), 1 (on), or 2 (mixed)")
    elif control.kind in {"Spin", "Scroll"}:
        if not control.minimum <= value <= numeric_maximum(control):
            raise ValueError("control value is outside its bounds")
    elif not 0 <= value <= list_count(sheet, shape):
        raise ValueError("list value must be zero or a one-based index within the list")
    if control.kind == "List" and control.selection_mode != "single":
        control.selected_indices = [value] if value else []
        sheet.drawing_changed()
        return
    target = linked_target(sheet, control.linked_cell)
    refresh(sheet, shape)
    wanted = -4146 if control.kind == "CheckBox" and value == 0 else value
    if wanted == control.value:
        return
    control.value = wanted
    sheet.drawing_changed()
    if target:
        other, row, column = target
        cell = other.cell(row, column, create=True)
        assert cell is not None
        cell.value = linked_value(control)
        cell.formula = ""
        cell.stale = False
        other.cell_changed(row, column)


def set_link(sheet: Worksheet, shape: Shape, reference: str) -> None:
    """VBA rebinding writes the control's state if the target value differs."""
    from pyopenvba.apps.excel._control_refs import binding

    reference, _ = binding(sheet, reference)
    if shape.control is not None and shape.control.kind == "Radio":
        from pyopenvba.apps.excel._radios import set_link as set_radio_link

        set_radio_link(sheet, shape, reference)
        return
    if shape.control is not None and shape.control.kind in {"Drop", "List"}:
        list_control(shape)
    else:
        value_control(shape)
    target = linked_target(sheet, reference)
    control = shape.control
    assert control is not None
    refresh(sheet, shape)
    wanted = linked_value(control)
    old: object = EMPTY
    if target:
        other, row, column = target
        old = other.book.calculator.value_of(other.name, row, column)
    control.linked_cell = reference
    sheet.drawing_changed()
    if target:
        other, row, column = target
        comparable = False if old is EMPTY else old
        if comparable != wanted:
            cell = other.cell(row, column, create=True)
            assert cell is not None
            cell.value, cell.formula, cell.stale = wanted, "", False
            other.cell_changed(row, column)


def numeric_control(shape: Shape) -> ControlInfo:
    control = shape.control
    if control is None or control.kind not in {"Spin", "Scroll"}:
        raise VBAUnsupportedError("numeric bounds require a spinner or scroll bar")
    return control


def numeric_maximum(control: ControlInfo) -> int:
    return control.maximum if control.maximum is not None else (30000 if control.kind == "Spin" else 100)


def set_numeric_property(sheet: Worksheet, shape: Shape, name: str, value: int) -> None:
    control = numeric_control(shape)
    if not 0 <= value <= 30000:
        raise ValueError("numeric control properties must be between 0 and 30000")
    target = linked_target(sheet, control.linked_cell)
    refresh(sheet, shape)
    setattr(control, name, value)
    maximum = numeric_maximum(control)
    # Excel accepts inverted bounds. Its link is lower-bound constrained,
    # while the displayed control value is upper-bound constrained.
    linked = max(control.minimum, min(control.value, maximum))
    control.value = min(linked, maximum)
    sheet.drawing_changed()
    if target:
        other, row, column = target
        cell = other.cell(row, column, create=True)
        assert cell is not None
        cell.value, cell.formula, cell.stale = float(linked), "", False
        other.cell_changed(row, column)


def cell_changed(sheet: Worksheet, row: int, column: int, value: object) -> None:
    """Update supported controls bound to this cell, including other worksheets."""
    if value is EMPTY:
        wanted = -4146
    elif isinstance(value, (bool, int, float)):
        wanted = 1 if value else -4146
    elif isinstance(value, ExcelError) and value.name == "#N/A":
        wanted = 2
    elif isinstance(value, VBAErrorValue) and value.number == 2042:
        wanted = 2
    else:
        # Other errors and text preserve checkbox state in Excel.
        wanted = None
    for owner in sheet.book.sheets_:
        from pyopenvba.apps.excel._radios import groups, cell_value

        for items in groups(owner):
            leader = items[0].control
            assert leader is not None
            try:
                if linked_target(owner, leader.linked_cell) == (sheet, row, column):
                    cell_value(owner, items, value)
            except VBAUnsupportedError:
                continue
        for shape in owner.shapes_:
            control = shape.control
            if control is None or control.kind not in {"CheckBox", "Drop", "List", "Spin", "Scroll"} or not control.linked_cell:
                continue
            try:
                target = linked_target(owner, control.linked_cell)
                selected = wanted
                if control.kind in {"Drop", "List"}:
                    if control.selection_mode != "single":
                        continue
                    selected = list_selection(value, list_count(owner, shape))
                elif control.kind in {"Spin", "Scroll"}:
                    selected = list_selection(value, numeric_maximum(control))
                    if selected is not None:
                        selected = min(max(selected, control.minimum), numeric_maximum(control))
            except VBAUnsupportedError:
                continue
            if target == (sheet, row, column) and selected is not None and control.value != selected:
                control.value = selected
                owner.drawing_changed()
