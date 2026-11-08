"""Shared name validation, scope and reference rewriting."""
from __future__ import annotations
import bisect
import re
import string
from functools import lru_cache
from typing import TYPE_CHECKING, Final
from pyopenvba._a1 import parse_area, quote_sheet
from pyopenvba.apps.excel._name_characters import REFUSED_FIRST, REFUSED_LATER
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
    if not valid(bare):
        raise error(1004, SYNTAX)
    return f"{quote_sheet(owner)}!{bare}" if owner else bare


#: The start of a name read as an R1C1 reference: R, C, RC, R1, C1, R1C1 and the like.
_R1C1_START = re.compile(r"(R([0-9]*))?(C([0-9]*))?", re.I)


def refers_text(value: object, empty: str = "=") -> str:
    """What setting a name's RefersTo to ``value`` holds: a formula as it is if it parses, and error 1004 if not; TRUE for
    True, other text quoted, which reads as text, and "=" for "", which is nothing at all (tests/fixtures/settable_lists.json). Null names the active cell, in
    the table's own words if it is in one, which reports itself."""
    from pyopenvba.exceptions import VBAUnsupportedError
    from pyopenvba.formula._parse import FormulaError, parse
    from pyopenvba.interpreter._values import EMPTY, NULL, to_text

    if value is NULL:
        raise VBAUnsupportedError("a name's RefersTo set to Null is not implemented")
    if value is EMPTY:
        return empty
    if isinstance(value, bool):
        return "=TRUE" if value else "=FALSE"
    text = to_text(value)
    if not isinstance(value, str):
        return "=" + text
    if not text:
        return "="
    if text == "=":
        return '="="'
    if not text.startswith("="):
        return '="' + text.replace('"', '""') + '"'
    try:
        parse(text)
    except FormulaError as failed:
        raise error(1004, "Application-defined or object-defined error") from failed
    return text


def _unit_ok(unit: int, first: bool) -> bool:
    """Whether Excel takes one UTF-16 unit of a name, first or after the first."""
    if unit < 0x80:
        return chr(unit) in (_FIRST_ASCII if first else _LATER_ASCII)
    refused = REFUSED_FIRST if first else REFUSED_LATER
    at = bisect.bisect_right(refused, (unit, 0x10000)) - 1
    return not (at >= 0 and refused[at][0] <= unit <= refused[at][1])


_FIRST_ASCII = frozenset(string.ascii_letters + "_\\")
_LATER_ASCII = frozenset(string.ascii_letters + string.digits + "_.\\?")
#: The zero of each run of digits Excel reads as ASCII digits in a reference; not the fullwidth digits, nor later scripts'.
_ADDRESS_ZEROS = (0x0660, 0x06F0, 0x0966, 0x09E6, 0x0A66, 0x0AE6, 0x0B66, 0x0BE6, 0x0C66, 0x0CE6, 0x0D66, 0x0E50, 0x0ED0, 0x0F20, 0x1040,
                  0x17E0, 0x1810)
_ADDRESS_DIGITS = {zero + digit: str(digit) for zero in _ADDRESS_ZEROS for digit in range(10)}


def valid(bare: str) -> bool:
    """Whether Excel takes ``bare`` as a name (tests/fixtures/name_validity.json): in ASCII a letter, _ or \\ first, then
    letters, digits, _, ., \\ and ?, at most 255 of them; no cell's address, A1 or RX1; and nothing that starts as
    an R1C1 reference -- R, C, RC or Rc alone, nor R1, C12 or RC1 with anything after it but a . -- though R_, Cx
    and Rb12c are names. Beyond ASCII, each UTF-16 unit but those _name_characters lists, which Excel refuses first or after
    the first (tests/fixtures/name_characters_bmp.json, name_surrogate_pairs.json)."""
    data = bare.encode("utf-16-le", "surrogatepass")
    units = [int.from_bytes(data[at:at + 2], "little") for at in range(0, len(data), 2)]
    if not units or len(bare) > 255:
        return False
    if not all(_unit_ok(unit, at == 0) for at, unit in enumerate(units)):
        return False
    if bare.upper() in ("TRUE", "FALSE"):
        return False
    # Read as a reference, the digits of seventeen scripts stand for their ASCII digits (tests/fixtures/name_address_digits.json,
    # name_address_digit_forms.json).
    bare = bare.translate(_ADDRESS_DIGITS)
    found = _R1C1_START.match(bare)
    if found is not None and found.group(0):
        if found.end() == len(bare):
            return False
        numbers = [number for number in (found.group(2), found.group(4)) if number]
        if numbers and not (bare[found.end()] == "." and len(numbers) == 1 and found.group(0)[-1].isdigit()):
            return False
    try:
        parse_area(bare)
    except ValueError:
        return True
    return False


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
