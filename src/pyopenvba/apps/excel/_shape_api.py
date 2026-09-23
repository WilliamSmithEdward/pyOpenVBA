"""Python worksheet shape operations, using the VBA model's state."""

from __future__ import annotations

import math
from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._shapes import FORM_CONTROLS, ShapeObject, Shapes
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.shapes import PRESET_GEOMETRY, Shape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet


def lookup(sheet: Worksheet, name: str) -> Shape:
    for item in sheet.shapes_:
        if item.name.lower() == name.lower():
            return item
    raise KeyError(f"there is no shape called {name!r} on {sheet.name}")


def snapshot(shape: Shape) -> Shape:
    copy = deepcopy(shape)
    if copy.control is not None:
        copy.control.macro_text = f"[0]!{copy.macro}" if copy.macro else ""
    return copy


def _validate(
    sheet: Worksheet, name: str | None, box: tuple[float, float, float, float],
    current: Shape | None = None,
) -> None:
    if name is not None:
        if not name.strip():
            raise ValueError("a shape name cannot be empty")
        if any(item is not current and item.name.lower() == name.lower() for item in sheet.shapes_):
            raise ValueError(f"a shape called {name!r} already exists on {sheet.name}")
    if any(not math.isfinite(value) or value < 0 for value in box):
        raise ValueError("shape coordinates and sizes must be finite, nonnegative points")


def validate_text(*values: str | None) -> None:
    for value in values:
        if value is None:
            continue
        for character in value:
            code = ord(character)
            if character not in "\t\n\r" and not (
                0x20 <= code <= 0xD7FF or 0xE000 <= code <= 0xFFFD or 0x10000 <= code <= 0x10FFFF
            ):
                raise ValueError("shape text, names and macros must contain valid XML characters")


def add(
    sheet: Worksheet, kind: str, shape_type: int, *, name: str | None,
    left: float, top: float, width: float, height: float, text: str, macro: str,
) -> Shape:
    _validate(sheet, name, (left, top, width, height))
    validate_text(name, text, macro)
    if kind == "shape" and shape_type not in PRESET_GEOMETRY:
        raise VBAUnsupportedError(f"AutoShape type {shape_type} has no measured writer")
    if kind == "formControl" and shape_type not in FORM_CONTROLS:
        raise ValueError(f"unsupported form control type {shape_type}")
    collection = Shapes(sheet)
    if kind == "formControl":
        made = collection.AddFormControl(shape_type, left, top, width, height)
    elif kind == "textBox":
        made = collection.AddTextbox(1, left, top, width, height)
    else:
        made = collection.AddShape(shape_type, left, top, width, height)
    assert isinstance(made, ShapeObject)
    item = made.shape
    if name is not None:
        item.name = name
    else:
        # The Python API offers unambiguous names. VBA itself can
        # generate a duplicate after a manual rename (measured in Excel).
        taken = {one.name.lower() for one in sheet.shapes_ if one is not item}
        stem, suffix = item.name, 2
        while item.name.lower() in taken:
            item.name = f"{stem} ({suffix})"
            suffix += 1
    item.text = text
    item.macro = macro
    return snapshot(item)


def update(
    sheet: Worksheet, name: str, *, new_name: str | None = None,
    left: float | None = None, top: float | None = None,
    width: float | None = None, height: float | None = None,
    text: str | None = None, macro: str | None = None,
) -> Shape:
    item = lookup(sheet, name)
    box = (
        item.left if left is None else left, item.top if top is None else top,
        item.width if width is None else width, item.height if height is None else height,
    )
    _validate(sheet, new_name, box, item)
    validate_text(new_name, text, macro)
    if item.kind not in {"shape", "textBox", "line", "formControl"}:
        raise VBAUnsupportedError(f"the Python shape editor does not yet edit {item.kind}")
    if text is not None and item.kind == "line":
        raise VBAUnsupportedError("a line has no editable text body")
    if new_name is not None:
        item.name = new_name
    item.left, item.top, item.width, item.height = box
    if text is not None:
        item.text = text
    if macro is not None:
        item.macro = macro
    sheet.drawing_changed()
    return snapshot(item)


def remove(sheet: Worksheet, name: str) -> None:
    item = lookup(sheet, name)
    if item.kind not in {"shape", "textBox", "line", "formControl", "group"}:
        raise VBAUnsupportedError(f"the Python shape editor does not yet remove {item.kind}")
    ShapeObject(sheet, item).Delete()


def update_control(
    sheet: Worksheet, name: str, *, linked_cell: str | None, list_range: str | None,
) -> Shape:
    return update_control_shape(sheet, lookup(sheet, name), linked_cell=linked_cell, list_range=list_range)


def update_control_shape(
    sheet: Worksheet, item: Shape, *, linked_cell: str | None, list_range: str | None,
) -> Shape:
    """Edit bindings and reconcile list selection without writing linked-cell values."""
    from pyopenvba.apps.excel._control_refs import binding

    control = item.control
    if item.kind != "formControl" or control is None:
        raise ValueError(f"{item.name!r} is not a form control")
    validate_text(linked_cell, list_range)
    if linked_cell is not None and control.kind not in {"CheckBox", "Drop", "List", "Radio", "Scroll", "Spin"}:
        raise VBAUnsupportedError(f"linked cells are not supported for {control.kind}")
    if list_range is not None and control.kind not in {"Drop", "List"}:
        raise VBAUnsupportedError(f"list ranges are not supported for {control.kind}")
    if linked_cell is not None:
        linked_cell, _ = binding(sheet, linked_cell)
    if list_range is not None:
        list_range, _ = binding(sheet, list_range)
    if linked_cell is not None:
        if control.kind == "Radio":
            from pyopenvba.apps.excel._radios import group

            leader = group(sheet, item)[0].control
            assert leader is not None
            leader.linked_cell = linked_cell
        else:
            control.linked_cell = linked_cell
    if list_range is not None:
        control.list_range = list_range
        if list_range:
            control.items.clear()
        from pyopenvba.apps.excel._controls import list_source_changed

        list_source_changed(sheet, item)
    sheet.drawing_changed()
    return snapshot(item)
