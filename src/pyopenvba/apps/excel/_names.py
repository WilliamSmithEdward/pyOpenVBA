"""Shared name validation, scope and reference rewriting."""
from __future__ import annotations
import re
from typing import TYPE_CHECKING
from pyopenvba._a1 import parse_area, quote_sheet
from pyopenvba.formula._parse import tokenize, split_sheet, FormulaError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, NameEntry, Worksheet


def canonical(book: Workbook, name: str, scope: Worksheet | None = None) -> str:
    owner, bare = split_sheet(name)
    workbook_qualified = owner.casefold() == book.name.casefold()
    if workbook_qualified:
        owner = ""
    if not owner and scope is not None and not workbook_qualified:
        owner = scope.name
    if owner:
        owner = book.sheet_named(owner).name
    if not re.fullmatch(r"[A-Za-z_\\\u0080-\uffff][\w.\\]*", bare) or len(bare) > 255 or bare.upper() in {"R", "C"}:
        raise error(1004, "Invalid defined name")
    if re.fullmatch(r"R[0-9]*C[0-9]*", bare, re.I):
        raise error(1004, "A name cannot be a cell reference")
    try:
        parse_area(bare)
    except ValueError:
        pass
    else:
        raise error(1004, "A name cannot be a cell reference")
    return f"{quote_sheet(owner)}!{bare}" if owner else bare


def qualify(formula: str, sheet: str) -> str:
    text = formula if formula.startswith("=") else "=" + formula
    try:
        tokens = tokenize(text)
    except FormulaError as exc:
        raise error(1004, str(exc)) from None
    if len(tokens) <= 2:
        raise error(1004, "A name requires a reference or formula")
    for token in reversed(tokens):
        if token.kind == "ref" and not split_sheet(token.text)[0]:
            text = text[:token.at] + quote_sheet(sheet) + "!" + token.text + text[token.at + len(token.text):]
    return text


def structured(formula: str) -> bool:
    """Whether a name's formula holds a structured reference, Table1[Qty]."""
    try:
        return any(token.kind == "structured" for token in tokenize(formula.removeprefix("=")))
    except FormulaError:
        return False


def renamed_formula(formula: str, owner: Worksheet, entry: NameEntry, wanted: str) -> str:
    text = formula
    new_scope, bare = split_sheet(wanted)
    for token in reversed(tokenize(formula)):
        if token.kind != "name":
            continue
        found = owner.book.names_.find(token.text, scope=owner)
        if found is not None and found.entry is entry:
            scope, _ = split_sheet(token.text)
            replacement = wanted if scope or (new_scope and new_scope.casefold() != owner.name.casefold()) else bare
            text = text[:token.at] + replacement + text[token.at + len(token.text):]
    return text


def changed(book: Workbook) -> None:
    book.names_.changed = True
    book.saved = False
    book.calculator.rebuild()



def owned_reserved(sheet: Worksheet, supported: tuple[str, ...]) -> list[tuple[NameEntry, str]]:
    """Select local reserved names with the exact metadata owned by these APIs."""
    from pyopenvba.apps.excel._advanced_filter import attribute_values

    result: list[tuple[NameEntry, str]] = []
    for entry in sheet.book.names_.entries:
        owner, bare = split_sheet(entry.name)
        if (owner.casefold() == sheet.name.casefold() and bare in supported
                and attribute_values(entry.attributes).get('name') == '_xlnm.' + bare):
            result.append((entry, bare))
    return result


def rename_owned_reserved(sheet: Worksheet, wanted: str) -> None:
    """Stage all supported reserved-name changes before changing any entry."""
    from pyopenvba.apps.excel._sheet_copy import formula
    from pyopenvba.exceptions import VBAUnsupportedError

    planned: list[tuple[NameEntry, str, str]] = []
    selected = owned_reserved(sheet, ('Print_Area', 'Print_Titles', '_FilterDatabase', 'Criteria', 'Extract'))
    seen: set[str] = set()
    for entry, bare in selected:
        name = quote_sheet(wanted) + '!' + bare
        if bare in seen or any(other is not entry and other.name.casefold() == name.casefold()
                               for other in sheet.book.names_.entries):
            raise VBAUnsupportedError('Renaming a worksheet with conflicting owned reserved names is not implemented')
        seen.add(bare)
        try:
            refers = formula(entry.refers_to, sheet, wanted, False)
        except FormulaError:
            raise VBAUnsupportedError('Renaming a worksheet with an unrecognized owned reserved name is not implemented') from None
        planned.append((entry, name, refers))
    for entry, name, refers in planned:
        entry.name, entry.refers_to = name, refers
    if planned:
        sheet.book.names_.changed = True
