"""Resolve Forms bindings without replacing their saved name spelling."""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_reference, split_sheet, quote_sheet
from pyopenvba._xml import attributes
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.formula._engine import Context
from pyopenvba.formula._values import ExcelError, as_bool, as_number, as_text, single

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet, NameEntry


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


def _formula_area(sheet: Worksheet, node: P.Node, seen: frozenset[str]) -> Area | None:
    """Resolve measured reference formulas without materializing their cells."""
    context = Context(sheet.book.calculator, sheet.name)
    if isinstance(node, P.Reference):
        return context.resolve(node)
    if isinstance(node, P.NameNode):
        text = f"{quote_sheet(node.sheet)}!{node.name}" if node.sheet else node.name
        return binding(sheet, text, seen=seen)[1]
    if isinstance(node, (P.Literal, P.Binary, P.Unary, P.ArrayLiteral)):
        return None
    if not isinstance(node, P.Call):
        raise VBAUnsupportedError("this control name does not return a supported reference")
    args = node.args
    if node.name.upper() == "CHOOSE" and len(args) >= 2:
        index = int(as_number(single(context.value(args[0]))))
        return _formula_area(sheet, args[index], seen) if 1 <= index < len(args) else None
    if node.name.upper() == "IF" and 2 <= len(args) <= 3:
        index = 1 if as_bool(single(context.value(args[0]))) else 2
        return _formula_area(sheet, args[index], seen) if index < len(args) else None
    if node.name.upper() == "INDEX" and 2 <= len(args) <= 4:
        area = _formula_area(sheet, args[0], seen)
        if area is None:
            return None
        row = int(as_number(single(context.value(args[1]))))
        column = int(as_number(single(context.value(args[2])))) if len(args) > 2 else 1
        if len(args) == 4 and int(as_number(single(context.value(args[3])))) != 1:
            return None
        if len(args) == 2:
            if area.rows == 1:
                row, column = 1, row
            elif area.columns != 1:
                return None
        if row < 0 or column < 0 or row > area.rows or column > area.columns:
            return None
        top = area.top + row - 1 if row else area.top
        left = area.left + column - 1 if column else area.left
        bottom = top if row else area.bottom
        right = left if column else area.right
        return Area(top, left, bottom, right, area.sheet)
    if node.name.upper() == "INDIRECT" and 1 <= len(args) <= 2:
        text = as_text(single(context.value(args[0])))
        if len(args) == 2 and not as_bool(single(context.value(args[1]))):
            if "[" in text or "]" in text:
                raise VBAUnsupportedError("relative or external R1C1 control names are not implemented")
            scope, address = split_sheet(text)
            match = re.fullmatch(r"R([0-9]+)C([0-9]+)(?::R([0-9]+)C([0-9]+))?", address, re.I)
            if match is None:
                return None
            first_row, first_column = int(match[1]), int(match[2])
            last_row, last_column = int(match[3] or match[1]), int(match[4] or match[2])
            top, bottom = sorted((first_row, last_row))
            left, right = sorted((first_column, last_column))
            if top < 1 or left < 1 or bottom > MAX_ROWS or right > MAX_COLUMNS:
                return None
            target = next((one for one in sheet.book.sheets_ if one.name.lower() == (scope or sheet.name).lower()), None)
            return Area(top, left, bottom, right, target.name) if target else None
        return binding(sheet, text, seen=seen)[1]
    if node.name.upper() == "OFFSET" and 3 <= len(args) <= 5:
        area = _formula_area(sheet, args[0], seen)
        if area is None:
            return None
        def number(index: int, default: int) -> int:
            if index >= len(args):
                return default
            argument = args[index]
            if isinstance(argument, P.Literal) and argument.value is None:
                return default
            return int(as_number(single(context.value(argument))))
        top = area.top + number(1, 0)
        left = area.left + number(2, 0)
        height, width = number(3, area.rows), number(4, area.columns)
        if top < 1 or left < 1 or height < 1 or width < 1:
            return None
        if top + height - 1 > MAX_ROWS or left + width - 1 > MAX_COLUMNS:
            return None
        return Area(top, left, top + height - 1, left + width - 1, area.sheet)
    raise VBAUnsupportedError("control reference formulas currently support CHOOSE, IF, INDEX, OFFSET and INDIRECT")


def binding(sheet: Worksheet, reference: str, *, single: bool = False,
            seen: frozenset[str] = frozenset()) -> tuple[str, Area | None]:
    text = reference.strip().removeprefix("=").strip()
    if not text:
        return "", None
    if "[" in text or "]" in text:
        raise VBAUnsupportedError("external workbook control references are not implemented")
    if re.match(r"[A-Za-z_][\w.]*\s*\(", text):
        try:
            return text, _formula_area(sheet, P.parse(text), seen)
        except ExcelError:
            return text, None
    scope, name = split_sheet(text)
    if scope and not any(one.name.lower() == scope.lower() for one in sheet.book.sheets_):
        raise ValueError(f"there is no worksheet called {scope!r}")
    try:
        areas = parse_reference(text)
    except ValueError:
        if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.\\]*", name):
            raise VBAUnsupportedError("control names must refer to an A1 range or another name") from None
        found = _name(sheet, scope, name)
        if found is None:
            return text, None
        entry, bare = found
        key = entry.name.lower() + "|" + entry.attributes
        if key in seen:
            raise VBAUnsupportedError("circular control names are not implemented")
        _, area = binding(sheet, entry.refers_to, seen=seen | {key})
        canonical = f"{quote_sheet(scope)}!{bare}" if scope else bare
        return canonical, area
    if len(areas) != 1 or (single and (areas[0].rows != 1 or areas[0].columns != 1)):
        raise ValueError("a linked cell must be one cell; a list range must be one rectangle")
    return text, areas[0]
