"""Resolve Forms bindings without replacing their saved name spelling."""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_reference, split_sheet, quote_sheet
from pyopenvba._xml import attributes
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.nodes import AreaReference, AxisReference, Binary, CellReference, NameReference, Node

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet, NameEntry

#: A defined name, as a control's link or list may give one.
_NAME = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]*")
#: Text in quotes inside a formula, where brackets are only text.
_TEXT = re.compile(r'"(?:[^"]|"")*"')
#: A reference into another workbook: [Book2]Sheet1!A1, or [1]!Name.
_EXTERNAL = re.compile(r"\][^\[\]!\"]*!")


def _name(sheet: Worksheet, scope: str, name: str) -> tuple[NameEntry, str] | None:
    local: tuple[NameEntry, str] | None = None
    global_name: tuple[NameEntry, str] | None = None
    for entry in sheet.book.names_.entries:
        owner, bare = split_sheet(entry.name)
        stored = attributes(f"<definedName {entry.attributes}>")
        local_id = stored.get("localSheetId")
        if not owner and local_id is not None and local_id.isdigit():
            index = int(local_id)
            if index < len(sheet.book.sheets_):
                owner = sheet.book.sheets_[index].name
        if bare.lower() != name.lower():
            continue
        if owner.lower() == (scope or sheet.name).lower():
            local = entry, bare
        elif not owner:
            global_name = entry, bare
    return local or global_name


def _written(node: Node | None) -> bool:
    """Whether a link or list is written as cells or a name, or references joined by commas, rather than as a
    formula; text the engine cannot read is taken as written."""
    if node is None or isinstance(node, (CellReference, AreaReference, AxisReference, NameReference)):
        return True
    return isinstance(node, Binary) and node.op == "," and _written(node.left) and _written(node.right)


def _formula_area(sheet: Worksheet, node: Node) -> Area | None:
    """The one block a formula lands on, worked out by the cell engine as a defined name's formula is, from A1:
    CHOOSE, IF, INDEX, OFFSET, INDIRECT, LET and the reference operators as measured
    (tests/fixtures/shapes/control_formula_names.json). A value, an error or more than one block is None."""
    from pyopenvba.formula._calc.evaluator import Context
    from pyopenvba.formula._calc.values import ExcelError, Reference

    calculator = sheet.book.calculator
    now = calculator.now()
    context = Context(calculator.engine_book, sheet.name, 1, 1, array=True, today=now.date(), now=now)
    try:
        value = context.evaluate(node)
    except ExcelError:
        return None
    if not isinstance(value, Reference) or value.area is None:
        return None
    found = value.area
    return Area(found.top, found.left, found.bottom, found.right, found.sheet)


def binding(sheet: Worksheet, reference: str, *, seen: frozenset[str] = frozenset()) -> tuple[str, Area | None]:
    """A link or list source as the control keeps it, and the one block of cells it lands on, or None for none.

    A link to more than one cell is the top left cell's, as Excel takes $H$1:$I$1; a name that comes back to itself
    lands on no cells (tests/fixtures/shapes/control_formula_names.json)."""
    text = reference.strip().removeprefix("=").strip()
    if not text:
        return "", None
    if _EXTERNAL.search(_TEXT.sub('""', text)):
        raise VBAUnsupportedError("external workbook control references are not implemented")
    try:
        node: Node | None = sheet.book.calculator.engine_book.read("=" + text)
    except FormulaSyntaxError:
        node = None
    if node is not None and not _written(node):
        return text, _formula_area(sheet, node)
    scope, name = split_sheet(text)
    if scope and not any(one.name.lower() == scope.lower() for one in sheet.book.sheets_):
        raise ValueError(f"there is no worksheet called {scope!r}")
    try:
        areas = parse_reference(text)
    except ValueError:
        if not _NAME.fullmatch(name):
            raise VBAUnsupportedError("control names must refer to an A1 range or another name") from None
        found = _name(sheet, scope, name)
        if found is None:
            return text, None
        entry, bare = found
        canonical = f"{quote_sheet(scope)}!{bare}" if scope else bare
        key = entry.name.lower() + "|" + entry.attributes
        if key in seen:
            return canonical, None
        _, area = binding(sheet, entry.refers_to, seen=seen | {key})
        return canonical, area
    if len(areas) != 1:
        raise ValueError("a link or a list range must be one rectangle")
    return text, areas[0]
