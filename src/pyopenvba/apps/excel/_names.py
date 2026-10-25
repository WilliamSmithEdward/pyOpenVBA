"""Shared name validation, scope and reference rewriting."""
from __future__ import annotations
import re
from functools import lru_cache
from typing import TYPE_CHECKING, Final
from pyopenvba._a1 import parse_area, quote_sheet
from pyopenvba.formula._parse import tokenize, split_sheet, FormulaError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, NameEntry, Worksheet

#: What Excel says of a name it will not take, a cell's address among them (live probe).
SYNTAX: Final = ("The syntax of this name isn't correct.\n\nVerify that the name:\n\t-Starts with a letter or underscore "
                 "(_)\n\t-Doesn't include a space or character that isn't allowed\n\t-Doesn't conflict with an existing "
                 "name in the workbook.")
#: And of a formula it will not take.
FORMULA_PROBLEM: Final = (
    "There's a problem with this formula.\n\nNot trying to type a formula?\nWhen the first character is an equal "
    "(\"=\") or minus (\"-\") sign, Excel thinks it's a formula:\n\n\u2022 you type:   =1+1, cell shows:   2\n\nTo get "
    "around this, type an apostrophe ( ' ) first:\n\n\u2022 you type:   '=1+1, cell shows:   =1+1")


def canonical(book: Workbook, name: str, scope: Worksheet | None = None) -> str:
    owner, bare = split_sheet(name)
    workbook_qualified = owner.casefold() == book.name.casefold()
    if workbook_qualified:
        owner = ""
    if not owner and scope is not None and not workbook_qualified:
        owner = scope.name
    if owner:
        owner = book.sheet_named(owner).name
    # After its first character a name may hold a combining accent, e with U+0301
    # (tests/fixtures/name_equality.json, marks).
    if (not re.fullmatch(r"[A-Za-z_\\\u0080-\uffff][\w.\\\u0300-\u036f]*", bare) or len(bare) > 255
            or bare.upper() in {"R", "C"}):
        raise error(1004, SYNTAX)
    if re.fullmatch(r"R[0-9]*C[0-9]*", bare, re.I):
        raise error(1004, SYNTAX)
    try:
        parse_area(bare)
    except ValueError:
        pass
    else:
        raise error(1004, SYNTAX)
    return f"{quote_sheet(owner)}!{bare}" if owner else bare


def refused_structured(formula: str, context: Worksheet) -> str | None:
    """What Excel says of a name's A1 formula naming one column of a table on the sheet the name is read on, the
    active sheet's for the workbook's Names and a sheet's own for its: Table[Qty] is a problem with the formula and
    Table[@Qty] a name's syntax; None where Excel takes it, as it takes the whole table, its parts, a span of
    columns, the same from any other sheet and any of them in R1C1 (tests/fixtures/names_add/)."""
    tables = {table.name.casefold() for table in context.tables}
    try:
        tokens = tokenize(formula if formula.startswith("=") else "=" + formula)
    except FormulaError:
        return None
    for token in tokens:
        found = re.fullmatch(r"(?:[^!]*!)?([^\[]+)\[(@?)[^\[\]#@][^\[\]]*\]", token.text)
        if token.kind == "structured" and found is not None and found.group(1).casefold() in tables:
            return SYNTAX if found.group(2) else FORMULA_PROBLEM
    return None


def qualify(formula: str, sheet: str) -> str:
    text = formula if formula.startswith("=") else "=" + formula
    try:
        tokens = tokenize(text)
    except FormulaError as exc:
        raise error(1004, str(exc)) from None
    if len(tokens) <= 2:
        raise error(1004, "A name requires a reference or formula")
    for token in reversed(tokens):
        if token.kind != "ref":
            continue
        owner, reference = split_sheet(token.text)
        if not owner:
            replacement = quote_sheet(sheet) + "!" + token.text
        elif ":" in owner or "[" in token.text.rpartition("!")[0]:
            # A 3D span or another workbook's sheet keeps the spelling it was given.
            continue
        else:
            # A sheet's name is spelled as Excel spells it, quoted only where it has to be:
            # ='Data'!$C$2 reads back =Data!$C$2 (tests/fixtures/subtotals/, outside_formulas).
            replacement = quote_sheet(owner) + "!" + reference
        text = text[:token.at] + replacement + text[token.at + len(token.text):]
    return text


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


def changed(book: Workbook, *names: str) -> None:
    """A name changed: the workbook's names to write again, and the formulas naming ``names``, the names added,
    redefined, renamed or deleted, to work out again; a name's comment or visibility changes no formula
    (tests/fixtures/calculation_state.json, name_*)."""
    book.names_.changed = True
    book.saved = False
    if names:
        book.calculator.names_changed({split_sheet(name)[1].lower() for name in names})



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


@lru_cache(maxsize=4096)
def name_key(text: str) -> bytes | str:
    """What Excel tells two names apart by, case aside: word sort's key, under which ss is a sharp s and ae the
    ligature while an accent or a hyphen tells two apart (tests/fixtures/style_order.json,
    tests/fixtures/name_equality.json); for a name the key is not reproduced for, its characters in any case."""
    from pyopenvba.apps.excel._sort import text_key
    from pyopenvba.exceptions import VBAUnsupportedError

    try:
        return text_key(text, False)
    except VBAUnsupportedError:
        return text.casefold()


def same_name(one: str, other: str) -> bool:
    """Whether Excel finds a name by another: where word sort cannot tell them apart, case aside, so AESOP finds
    Aesop with the ash. Worksheets(...), Names(...), Names.Add and a formula's names find a name so
    (tests/fixtures/name_equality.json)."""
    return one == other or name_key(one) == name_key(other)


def same_sheet_name(one: str, other: str) -> bool:
    """Whether Excel takes two names for one where it keeps names apart: as long as each other, and word sort cannot
    tell them apart, so a Kelvin sign is a K and a final sigma a sigma while ae is not the ash. A sheet's name must
    differ so from every other, a formula's sheet is found so, and Names.Add takes the spelling it is given where
    the name's own is the same so (tests/fixtures/name_equality.json)."""
    return one == other or (len(one) == len(other) and name_key(one) == name_key(other))


def entry_key(name: str) -> tuple[int, bytes | str, bytes | str]:
    """What tells a defined name, with its sheet's name and a ! before it where it is one sheet's own, from every
    other: its sheet as Excel keeps sheets apart, and its name as word sort tells names apart."""
    scope, bare = split_sheet(name)
    return len(scope), name_key(scope), name_key(bare)


def same_entry(one: str, other: str) -> bool:
    """Whether two defined names, each with its sheet's name and a ! before it where it is one sheet's own, are one
    name (:func:`entry_key`)."""
    return one == other or entry_key(one) == entry_key(other)


def collation(text: str) -> bytes:
    """The name word ordering shared by iteration and persisted definitions: Windows' word sort, which puts
    eclair before its accented spelling as Excel does (tests/fixtures/document_properties/, names_more); text
    it does not reproduce sorts after the rest, by its characters."""
    from pyopenvba.apps.excel._sort import text_key
    from pyopenvba.exceptions import VBAUnsupportedError

    try:
        return text_key(text, False)
    except VBAUnsupportedError:
        return b"\xff" + text.lower().encode("utf-32-be")
