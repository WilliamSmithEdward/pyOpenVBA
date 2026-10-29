"""What a macro does with a workbook's links to other workbooks, and what a save writes of them.

LinkSources lists the links by the full names of the workbooks they read,
in the order the file keeps them. LinkInfo answers a link's status: whether
its workbook's file is there, and has every sheet the formulas read.
BreakLink turns each formula that reads a link, itself or through names,
into the value it shows, and points each name that reads the link itself at
#REF!; the link goes, the ones after it move up, and nothing is raised or
worked out again. A broken link is still known by its name to LinkInfo and
BreakLink until the workbook closes. UpdateLinks and SaveLinkValues are the
workbook's two link settings, kept in its workbookPr. Measured in live
Excel with every linked workbook closed, as the model keeps them
(scripts/measure_workbook_links.py, tests/fixtures/workbook_links.json).

A save writes the parts of the links again where Excel's save would differ
from the file read: a broken link's part goes and the rest are numbered
again; a link that also read cells only a broken formula read keeps the
values of the cells still read; SaveLinkValues False keeps no values; and a
link one of whose workbook's own names a formula reads gains a sheet named
after that workbook, as Excel adds when it saves a link whose workbook is
closed.
"""
from __future__ import annotations

import math
import re
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass, field, replace
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area, split_sheet
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.interpreter._values import (EMPTY, MISSING, NULL, VBAArray, VBAErrorValue, VBAInt, error, to_bool,
                                           to_integer, to_number, to_text, type_name)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._external_links import LinkedBook
    from pyopenvba.apps.excel._model import NameEntry, Workbook, Worksheet
    from pyopenvba.powerquery._opc import OpcFile

#: XlLinkType's xlLinkTypeExcelLinks and xlLinkTypeOLELinks; XlLinkInfo's xlLinkInfoStatus.
_EXCEL_LINKS, _OLE_LINKS, _STATUS = 1, 2, 3
#: The statuses of a link whose workbook is closed: xlLinkStatusMissingFile, xlLinkStatusMissingSheet and
#: xlLinkStatusOld, "status may be out of date".
_MISSING_FILE, _MISSING_SHEET, _OLD = 1, 2, 3
#: #VALUE! and #N/A, as LinkInfo answers them.
_VALUE, _NA = 2015, 2042
_BREAK_FAILED = "Method 'BreakLink' of object '_Workbook' failed"
#: A reference to a linked workbook as a file spells it: [1]Sheet1!A1, [1]Sheet1:Other!A1, '[1]My Sheet'!A1, [1]!Rate.
_REFERENCE = re.compile(r"(?:'\[(?P<quoted_number>[0-9]+)\](?P<quoted>(?:[^']|'')*)'|\[(?P<number>[0-9]+)\]"
                        r"(?P<sheets>[^'!\[\]]*))!(?P<rest>.*)", re.DOTALL)


@dataclass
class LinkChanges:
    """What a workbook's links went through since it was read, for its next save."""

    #: The parts of the links broken, which the save leaves out.
    broken: list[str] = field(default_factory=lambda: [])
    #: The parts of the links that a broken formula also read, whose values the save keeps for the cells still read.
    pruned: set[str] = field(default_factory=lambda: set())
    #: The parts a save with SaveLinkValues False wrote without their values, as they were with them: the workbook
    #: still has the values, and a save with it True again writes them.
    values: dict[str, str] = field(default_factory=lambda: {})


@dataclass(frozen=True)
class _Reference:
    """One reference to a linked workbook in a formula: the link's number, the sheets it names, first and last, and
    what follows the !, an area or one of the linked workbook's own names."""

    number: int
    first: str
    last: str
    rest: str
    #: Where the reference is in the formula's text, without its =, and how long it is.
    start: int
    length: int


def _references(formula: str) -> Iterator[_Reference]:
    """Each reference to a linked workbook in a formula as the model keeps it, = and all."""
    body = formula[1:] if formula.startswith("=") else formula
    if "[" not in body:
        return
    try:
        tokens = P.tokenize(body, spaces=True)
    except P.FormulaError as failure:
        raise VBAUnsupportedError(f"a formula the model cannot read, {formula}, among a workbook's links") from failure
    for token in tokens:
        if token.kind not in ("ref", "name"):
            continue
        found = _REFERENCE.fullmatch(token.text)
        if found is None:
            continue
        if found.group("quoted_number") is not None:
            number, sheets = found.group("quoted_number"), found.group("quoted").replace("''", "'")
        else:
            number, sheets = found.group("number"), found.group("sheets")
        first, _, last = sheets.partition(":")
        yield _Reference(int(number), first, last or first, found.group("rest"), token.at, len(token.text))


def _kind(value: object) -> int | None:
    """A link Type as Excel reads one: missing or Null xlExcelLinks; True 1 and False 0, as a worksheet counts them;
    a number cut to its whole part, 1.5 being 1; text the number it spells, or None where it spells none
    (tests/fixtures/workbook_links.json)."""
    if value is MISSING or value is NULL:
        return _EXCEL_LINKS
    if isinstance(value, bool):
        return int(value)
    if value is EMPTY:
        return 0
    if isinstance(value, str):
        try:
            value = to_number(value)
        except VBARuntimeError:
            return None
    if isinstance(value, (int, float, Decimal)):
        return math.trunc(value)
    raise VBAUnsupportedError(f"a link Type given as {type_name(value)} is not implemented")


def _links(book: Workbook) -> list[LinkedBook]:
    """The workbook's links to other workbooks, in the order the file keeps them."""
    return [book.linked_books[number] for number in sorted(book.linked_books)]


def sources(book: Workbook, kind: object) -> object:
    """Workbook.LinkSources: the full names of the workbooks the workbook links to, as an array from 1, or Empty
    where it has none; any Type but xlExcelLinks is Empty, the model keeping no other kind of link."""
    if _kind(kind) != _EXCEL_LINKS:
        if book.foreign_links:
            raise VBAUnsupportedError("LinkSources of a workbook with OLE or DDE links is not implemented")
        return EMPTY
    paths: list[object] = [linked.path(book.path) for linked in _links(book)]
    return VBAArray([(1, len(paths))], items=paths) if paths else EMPTY


def _found(book: Workbook, name: str) -> tuple[LinkedBook | None, bool]:
    """The link ``name`` gives, by the linked workbook's full name or its file's name alone, in either case and with
    / for \\; and whether it is one broken already."""
    wanted = name.replace("/", "\\").casefold()
    for broken, links in ((False, _links(book)), (True, book.broken_links)):
        for linked in links:
            if wanted in (linked.path(book.path).casefold(), linked.file.casefold()):
                return linked, broken
    return None, False


def _refuse_open(book: Workbook, linked: LinkedBook) -> None:
    """Report a link whose workbook is open too: Excel then works with that workbook, the model with the link's
    copy of its cells."""
    full = linked.path(book.path).casefold()
    for other in book.application.workbooks_.books:
        if other is not book and str(other.FullName()).casefold() == full:
            raise VBAUnsupportedError("a link to a workbook that is open is not implemented: the model reads the "
                                      "link's copy of its cells")


def info(book: Workbook, name: object, what: object, kind: object, edition: object) -> object:
    """Workbook.LinkInfo. Its name is String and its LinkInfo a Long, as VBA passes them; its Type a Variant.
    xlLinkInfoStatus of a link answers its status; any other LinkInfo, or a Type other than xlExcelLinks, is #N/A;
    a name that is no link #VALUE!, and an empty one #N/A (tests/fixtures/workbook_links.json)."""
    if name is MISSING or what is MISSING:
        raise error(449, "Argument not optional")
    text = to_text(name)
    asked = int(to_integer(what, "Long"))
    if edition is not MISSING:
        raise VBAUnsupportedError("LinkInfo's EditionRef, for the Mac's publishers and subscribers, is not implemented")
    which = _kind(kind)
    if which is None:
        return VBAErrorValue(_VALUE)
    if not text:
        return VBAErrorValue(_NA)
    linked, _ = _found(book, text)
    if linked is None:
        return VBAErrorValue(_VALUE)
    if which != _EXCEL_LINKS or asked != _STATUS:
        return VBAErrorValue(_NA)
    return VBAInt(status(book, linked), "Long")


def status(book: Workbook, linked: LinkedBook) -> int:
    """A link's status: its workbook's file not there, a sheet a formula or name reads not in it, or else the
    status of a link whose workbook is closed, which may be out of date. A sheet the link lists that nothing
    reads, or the one Excel adds for its workbook's names, is not missed (tests/fixtures/workbook_links.json)."""
    _refuse_open(book, linked)
    path = Path(linked.path(book.path))
    if not path.is_file():
        return _MISSING_FILE
    there = {sheet.casefold() for sheet in _sheet_names(path)}
    read = read_sheets(book, linked)
    return _MISSING_SHEET if any(sheet.casefold() not in there for sheet in read) else _OLD


def _sheet_names(path: Path) -> list[str]:
    """The sheets of the workbook in a file, from its workbook part."""
    from pyopenvba.apps.excel._xstring import decode

    try:
        with zipfile.ZipFile(path) as package:
            text = package.read("xl/workbook.xml").decode("utf-8", errors="replace")
    except (zipfile.BadZipFile, KeyError, OSError) as failure:
        raise VBAUnsupportedError(f"LinkInfo's status of a link to {path.name}, which is no xlsx, xlsm or xltx "
                                  "workbook, is not implemented") from failure
    from html import unescape

    return [decode(unescape(name)) for name in re.findall(r'<sheet\b[^>]*?\bname="([^"]*)"', text)]


def _formulas(book: Workbook) -> Iterator[tuple[str, Worksheet | None]]:
    """Every formula of the workbook's cells and names, with the sheet it is read on, None for a workbook's name."""
    from pyopenvba._a1 import split_sheet as split

    for sheet in book.sheets_:
        for cell in sheet.cells_.values():
            if cell.formula:
                yield cell.formula, sheet
    for entry in book.names_.entries:
        scope = split(entry.name)[0]
        yield entry.refers_to, book.sheet_named(scope) if scope else None


def read_sheets(book: Workbook, linked: LinkedBook) -> set[str]:
    """The sheets of a linked workbook the workbook's formulas and names read, the sheets between a 3-D reference's
    first and last too, and those a name of the linked workbook reads."""
    read: set[str] = set()
    order = [sheet.casefold() for sheet in linked.sheets]
    for formula, _ in _formulas(book):
        for reference in _references(formula):
            if reference.number != linked.number:
                continue
            if not reference.first:
                read.update(_name_sheets(linked, reference.rest))
                continue
            first, last = reference.first.casefold(), reference.last.casefold()
            if first in order and last in order:
                low, high = sorted((order.index(first), order.index(last)))
                read.update(linked.sheets[low:high + 1])
            else:
                read.update((reference.first, reference.last))
    return read


def _name_sheets(linked: LinkedBook, name: str) -> list[str]:
    """The sheets one of a linked workbook's own names reads, its formula as that workbook spells it: ='Sheet1'!$C$1."""
    refers = next((text for key, text in linked.names.items() if key.casefold() == name.casefold()), "")
    sheets: list[str] = []
    try:
        tokens = P.tokenize(refers.removeprefix("="), spaces=True) if refers else []
    except P.FormulaError:
        return sheets
    for token in tokens:
        if token.kind == "ref" and "!" in token.text:
            sheets.append(split_sheet(token.text)[0])
    return sheets


# -- BreakLink


def break_link(book: Workbook, name: object, kind: object) -> None:
    """Workbook.BreakLink. Its name is String and its Type a Long, as VBA passes them (tests/fixtures/workbook_links.json):
    a Type other than xlLinkTypeExcelLinks or xlLinkTypeOLELinks is error 5, and an OLE link, or a name that is no
    link, error 1004; a link broken already is left as it is."""
    if name is MISSING or kind is MISSING:
        raise error(450, "Wrong number of arguments or invalid property assignment")
    text = to_text(name)
    which = int(to_integer(kind, "Long"))
    if which not in (_EXCEL_LINKS, _OLE_LINKS):
        raise error(5, "Invalid procedure call or argument")
    linked, broken = _found(book, text) if which == _EXCEL_LINKS else (None, False)
    if linked is None:
        raise error(1004, _BREAK_FAILED)
    if broken:
        return
    _refuse_open(book, linked)
    if book.foreign_links:
        raise VBAUnsupportedError("BreakLink in a workbook with OLE or DDE links is not implemented")
    _refuse_elsewhere(book)
    frozen = _frozen(book, linked.number)
    also_read = {number for sheet, row, column in frozen
                 for number in _links_read(book, sheet.cells_[(row, column)].formula, sheet)} - {linked.number}
    values = {(sheet.name, row, column): book.calculator.value_of(sheet.name, row, column)
              for sheet, row, column in _with_blocks(frozen)}
    for sheet, row, column in _with_blocks(frozen):
        _freeze(sheet, row, column, values[(sheet.name, row, column)])
    for entry in book.names_.entries:
        pointed = _pointed_at_ref(entry.refers_to, linked.number)
        if pointed != entry.refers_to:
            entry.refers_to = pointed
            book.names_.changed = True
    changes = book.link_changes
    changes.broken.append(linked.part)
    changes.pruned.update(book.linked_books[number].part for number in also_read if number in book.linked_books)
    book.broken_links.append(linked)
    _renumber(book, linked.number)
    book.saved = False


def _refuse_elsewhere(book: Workbook) -> None:
    """Report a link a formula the model does not break reads: a validation's, a conditional format's, a chart's
    series, a control's cells, or a table's column."""
    found = re.compile(r"\[[0-9]+\]")
    for sheet in book.sheets_:
        texts: list[str] = [text for rule in sheet.validations for text in (rule.formula1, rule.formula2)]
        texts += ["".join(rule.root.itertext()) for rule in sheet.conditional_formats]
        texts += [text for chart in sheet.chart_states.values() for series in chart.series
                  for text in (series.name_reference, series.category_reference, series.value_reference)]
        texts += [text for shape in sheet.shapes_ if shape.control is not None
                  for text in (shape.control.linked_cell, shape.control.list_range)]
        if any(found.search(text) for text in texts if text):
            raise VBAUnsupportedError("BreakLink of a link a validation, conditional format, chart or control reads "
                                      "is not implemented")
    package = book.package
    if package is not None:
        for part in package.names():
            if re.fullmatch(r"xl/(charts/chart|tables/table)[0-9]+\.xml", part) \
                    and re.search(rb"<(?:c:f|calculatedColumnFormula)\b[^>]*>[^<]*\[[0-9]+\]", package.read(part)):
                raise VBAUnsupportedError("BreakLink of a link a chart or a table's column reads is not implemented")


def _links_read(book: Workbook, formula: str, sheet: Worksheet | None, seen: frozenset[int] = frozenset()) -> set[int]:
    """The links a formula on ``sheet`` reads, itself or through a name it names, and a name that one names
    (tests/fixtures/workbook_links.json: =Outer, Outer being =SUM(Ext)*2 and Ext Alpha's cells, reads Alpha)."""
    read = {reference.number for reference in _references(formula)}
    body = formula[1:] if formula.startswith("=") else formula
    try:
        tokens = P.tokenize(body, spaces=True)
    except P.FormulaError:
        return read
    for at, token in enumerate(tokens):
        if token.kind != "name" or token.text.startswith(("[", "'[")) \
                or at + 1 < len(tokens) and tokens[at + 1].kind == "open":
            continue
        entry = _name_entry(book, token.text, sheet)
        if entry is None or id(entry) in seen:
            continue
        owner = book.sheet_named(split_sheet(entry.name)[0]) if "!" in entry.name else None
        read |= _links_read(book, entry.refers_to, owner, seen | {id(entry)})
    return read


def _name_entry(book: Workbook, text: str, sheet: Worksheet | None) -> NameEntry | None:
    """The name a formula on ``sheet`` means by ``text``: Sheet1!Rate that sheet's, a bare one its sheet's own
    before the workbook's."""
    from pyopenvba.apps.excel._names import same_name, same_sheet_name

    scope, bare = split_sheet(text)
    candidates = [(entry, *split_sheet(entry.name)) for entry in book.names_.entries]
    if scope:
        return next((entry for entry, own, entry_bare in candidates
                     if own and same_sheet_name(own, scope) and same_name(entry_bare, bare)), None)
    if sheet is not None:
        found = next((entry for entry, own, entry_bare in candidates
                      if own and same_sheet_name(own, sheet.name) and same_name(entry_bare, bare)), None)
        if found is not None:
            return found
    return next((entry for entry, own, entry_bare in candidates if not own and same_name(entry_bare, bare)), None)


def _frozen(book: Workbook, number: int) -> list[tuple[Worksheet, int, int]]:
    """The cells whose formulas read link ``number``, which breaking it turns into values; a protected sheet's are
    reported, Excel stopping there with an alert DisplayAlerts does not hold back."""
    cells: list[tuple[Worksheet, int, int]] = []
    for sheet in book.sheets_:
        for (row, column), cell in sorted(sheet.cells_.items()):
            if cell.formula and number in _links_read(book, cell.formula, sheet):
                if sheet.protection is not None:
                    raise VBAUnsupportedError("BreakLink of a link a protected sheet's formulas read is not "
                                              "implemented: Excel stops with an alert DisplayAlerts does not hold back")
                cells.append((sheet, row, column))
    return cells


def _with_blocks(cells: list[tuple[Worksheet, int, int]]) -> list[tuple[Worksheet, int, int]]:
    """The cells, and every other cell of an array formula's block or a dynamic formula's spill they start."""
    out: list[tuple[Worksheet, int, int]] = []
    for sheet, row, column in cells:
        block = sheet.array_formulas.get((row, column))
        spill = sheet.spills.get((row, column))
        area = block if block is not None else spill.area if spill is not None and not spill.blocked else None
        if area is None:
            out.append((sheet, row, column))
            continue
        for one_row in range(area.top, area.bottom + 1):
            for one_column in range(area.left, area.right + 1):
                found = sheet.cells_.get((one_row, one_column))
                if block is not None or (one_row, one_column) == (row, column) \
                        or found is not None and found.spilled_from == (row, column):
                    out.append((sheet, one_row, one_column))
    return out


def _freeze(sheet: Worksheet, row: int, column: int, value: object) -> None:
    """One cell's formula, or its place in an array or a spill, turned into the value it shows, its format kept."""
    from pyopenvba.formula._values import BLANK

    sheet.array_formulas.pop((row, column), None)
    sheet.spills.pop((row, column), None)
    cell = sheet.cell(row, column, create=True)
    assert cell is not None
    cell.value = EMPTY if value is BLANK else value
    cell.formula, cell.stale, cell.shared = "", False, None
    cell.dynamic, cell.spilled_from, cell.save_dirty = False, None, None
    sheet.cell_changed(row, column)


def _pointed_at_ref(formula: str, number: int) -> str:
    """A name's formula with each reference to link ``number`` pointed at #REF!: [2]Sheet1!$A$1:$A$3 is
    #REF!$A$1:$A$3, and [2]!Rate #REF!."""
    out: list[str] = []
    at = 0
    body = formula[1:] if formula.startswith("=") else formula
    for reference in _references(formula):
        if reference.number != number:
            continue
        out.append(body[at:reference.start])
        out.append("#REF!" + (reference.rest if reference.first else ""))
        at = reference.start + reference.length
    if not out:
        return formula
    return ("=" if formula.startswith("=") else "") + "".join(out) + body[at:]


def _renumber(book: Workbook, gone: int) -> None:
    """Link ``gone`` taken out: each later link one number earlier, in the links and in every formula, which is read
    again but not worked out again."""
    def moved(formula: str) -> str:
        body = formula[1:] if formula.startswith("=") else formula
        out: list[str] = []
        at = 0
        for reference in _references(formula):
            if reference.number > gone:
                text = body[reference.start:reference.start + reference.length]
                out.append(body[at:reference.start])
                out.append(text.replace(f"[{reference.number}]", f"[{reference.number - 1}]", 1))
                at = reference.start + reference.length
        if not out:
            return formula
        return ("=" if formula.startswith("=") else "") + "".join(out) + body[at:]

    book.linked_books = {(number - 1 if number > gone else number): (
        replace(linked, number=number - 1) if number > gone else linked)
        for number, linked in book.linked_books.items() if number != gone}
    for sheet in book.sheets_:
        for (row, column), cell in sheet.cells_.items():
            if cell.formula and "[" in cell.formula:
                changed = moved(cell.formula)
                if changed != cell.formula:
                    cell.formula = changed
                    sheet.touched()
                    book.calculator.remember(sheet.name, row, column, cell)
    for entry in book.names_.entries:
        changed = moved(entry.refers_to)
        if changed != entry.refers_to:
            entry.refers_to = changed
            book.names_.changed = True


def read_name(formula: str) -> str:
    """A name's formula as Excel reads it from a file: each #REF! a reference whose sheet and cells are both gone,
    #REF!#REF!, so #REF!*2 is =#REF!#REF!*2 (tests/fixtures/workbook_links.json, after BreakLink and after a sheet
    is deleted)."""
    if "#REF!" not in formula:
        return formula
    body = formula[1:] if formula.startswith("=") else formula
    try:
        tokens = P.tokenize(body, spaces=True)
    except P.FormulaError:
        return formula
    read = "".join("#REF!#REF!" if token.kind == "error" and token.text == "#REF!" else token.text for token in tokens)
    return ("=" if formula.startswith("=") else "") + read


def saved_name(formula: str) -> str:
    """A name's formula, without its =, as Excel saves it: a reference whose workbook or sheet is gone,
    #REF!$A$1:$A$3 or #REF!#REF!, as #REF! (tests/fixtures/workbook_links.json, BrokeAlpha.xlsx and Gone.xlsx)."""
    if "#REF!" not in formula:
        return formula
    try:
        tokens = P.tokenize(formula, spaces=True)
    except P.FormulaError:
        return formula
    out: list[str] = []
    gone = False
    for token in tokens:
        if gone and (token.kind == "ref" or token.kind == "error" and token.text == "#REF!"):
            gone = False
            continue
        gone = token.kind == "error" and token.text == "#REF!"
        out.append(token.text)
    return "".join(out)


# -- the settings


def set_update_links(book: Workbook, value: object) -> None:
    """UpdateLinks set: xlUpdateLinksUserSetting, Never or Always, a Long as VBA passes it, else error 1004; the
    workbook stays saved (tests/fixtures/workbook_links.json)."""
    wanted = int(to_integer(value, "Long"))
    if wanted not in (1, 2, 3):
        raise error(1004, "Method 'UpdateLinks' of object '_Workbook' failed")
    book.update_links = wanted


def set_save_link_values(book: Workbook, value: object) -> None:
    """SaveLinkValues set, a Boolean as VBA passes it; the workbook is no longer saved, even set as it was."""
    book.save_link_values = to_bool(value)
    book.saved = False


# -- saving


def write(book: Workbook, package: OpcFile) -> None:
    """The links' parts as Excel saves them where they differ from those read (see the module's docstring), and
    the workbook's link settings in its workbookPr."""
    _write_settings(book, package)
    changes = book.link_changes
    if book.save_link_values:
        for part, text in changes.values.items():
            if package.has(part):
                package.write(part, text.encode("utf-8"))
        changes.values.clear()
    links = _links(book)
    phantom = {linked.part: linked for linked in links if _reads_book_names(book, linked.number)}
    if not changes.broken and not changes.pruned and book.save_link_values and not phantom \
            and all(linked.part == f"xl/externalLinks/externalLink{linked.number}.xml" or not linked.part
                    for linked in links):
        return
    _drop_parts(book, package, changes.broken)
    for part in changes.broken:
        changes.values.pop(part, None)
    for linked in links:
        if not linked.part or not package.has(linked.part):
            continue
        text = package.read(linked.part).decode("utf-8")
        if linked.part in changes.pruned:
            text = _pruned(book, linked, text)
        if linked.part in phantom:
            text = _with_book_sheet(linked, text)
        if not book.save_link_values:
            changes.values.setdefault(linked.part, text)
            text = re.sub(r"<sheetDataSet>.*?</sheetDataSet>|<sheetDataSet/>", "", text, count=1, flags=re.DOTALL)
        package.write(linked.part, text.encode("utf-8"))
    _number_parts(book, package)
    changes.broken.clear()
    changes.pruned.clear()


def _write_settings(book: Workbook, package: OpcFile) -> None:
    """UpdateLinks and SaveLinkValues in the workbookPr, in the schema's order, where they differ from the file's."""
    if not package.has("xl/workbook.xml"):
        return
    text = package.read("xl/workbook.xml").decode("utf-8")
    found = re.search(r"<workbookPr\b[^>]*?/?>", text)
    tag = found.group(0) if found is not None else "<workbookPr/>"
    wanted = {"saveExternalLinkValues": None if book.save_link_values else "0",
              "updateLinks": {2: "never", 3: "always"}.get(book.update_links)}
    attributes = re.findall(r'\s([\w:]+)="([^"]*)"', tag)
    have = dict(attributes)
    if all(have.get(key) == value or key not in have and value is None
           or key == "updateLinks" and value is None and have.get(key) == "userSet" for key, value in wanted.items()):
        return
    order = ["date1904", "showObjects", "showBorderUnselectedTables", "filterPrivacy", "promptedSolutions",
             "showInkAnnotation", "backupFile", "saveExternalLinkValues", "updateLinks", "codeName",
             "hidePivotFieldList", "showPivotChartFilter", "allowRefreshQuery", "publishItems", "checkCompatibility",
             "autoCompressPictures", "refreshAllConnections", "defaultThemeVersion"]
    kept = [(key, value) for key, value in attributes if key not in wanted]
    for key, value in wanted.items():
        if value is None:
            continue
        rank = order.index(key)
        at = next((index for index, (other, _) in enumerate(kept) if other in order and order.index(other) > rank),
                  len(kept))
        kept.insert(at, (key, value))
    rebuilt = "<workbookPr" + "".join(f' {key}="{value}"' for key, value in kept) + "/>"
    if found is not None:
        text = text[:found.start()] + rebuilt + text[found.end():]
    else:
        place = re.search(r"<fileVersion\b[^>]*/>|<fileSharing\b[^>]*/>", text)
        at = place.end() if place is not None else text.index(">", text.index("<workbook")) + 1
        text = text[:at] + rebuilt + text[at:]
    package.write("xl/workbook.xml", text.encode("utf-8"))


def _reads_book_names(book: Workbook, number: int) -> bool:
    """Whether a formula or name reads one of link ``number``'s workbook's own names: [1]!Rate."""
    return any(reference.number == number and not reference.first
               for formula, _ in _formulas(book) for reference in _references(formula))


def _with_book_sheet(linked: LinkedBook, text: str) -> str:
    """A link's part with the sheet Excel adds, named after its workbook, for its workbook's names: listed last,
    with an empty sheetData marked refreshError (tests/fixtures/workbook_links.json, NamedAgain.xlsx)."""
    from html import escape, unescape

    from pyopenvba.apps.excel._xstring import decode, encode_attribute

    name = linked.file.rpartition(".")[0] or linked.file
    listed = [decode(unescape(sheet)) for sheet in re.findall(r'<sheetName\b[^>]*\bval="([^"]*)"', text)]
    if any(sheet.casefold() == name.casefold() for sheet in listed):
        return text
    escaped = escape(encode_attribute(name))
    text = text.replace("</sheetNames>", f'<sheetName val="{escaped}"/></sheetNames>', 1)
    return text.replace("</sheetDataSet>", f'<sheetData sheetId="{len(listed)}" refreshError="1"/></sheetDataSet>', 1)


def _pruned(book: Workbook, linked: LinkedBook, text: str) -> str:
    """A link's part keeping the values of the cells the workbook still reads and no others
    (tests/fixtures/workbook_links.json, MixedAgain.xlsx); as read where a reference cannot be followed."""
    areas = _read_areas(book, linked)
    if areas is None:
        return text

    def sheet_data(found: re.Match[str]) -> str:
        if found.group("empty") is not None:
            return found.group(0)
        index = int(found.group("index"))
        kept_areas = areas.get(index, [])
        rows: list[str] = []
        for row in re.finditer(r"<row\b[^>]*>(.*?)</row>", found.group("rows"), re.DOTALL):
            cells = [cell.group(0) for cell in re.finditer(r'<cell\b[^>]*?\br="([A-Z]+)([0-9]+)"[^>]*?(?:/>|>.*?</cell>)',
                                                         row.group(1), re.DOTALL)
                     if any(area.contains(int(cell.group(2)), _column(cell.group(1))) for area in kept_areas)]
            if cells:
                rows.append(row.group(0)[:row.group(0).index(">") + 1] + "".join(cells) + "</row>")
        opening = found.group("opening")
        return opening + "".join(rows) + "</sheetData>" if rows else opening[:-1] + "/>"

    return re.sub(r'(?P<empty><sheetData\b[^>]*/>)|(?P<opening><sheetData\b[^>]*?\bsheetId="(?P<index>[0-9]+)"[^>]*>)'
                  r'(?P<rows>.*?)</sheetData>', sheet_data, text, flags=re.DOTALL)


def _column(letters: str) -> int:
    from pyopenvba._a1 import column_number

    return column_number(letters)


def _read_areas(book: Workbook, linked: LinkedBook) -> dict[int, list[Area]] | None:
    """The areas of each of a linked workbook's sheets, by its place in the link, the workbook's formulas and names
    read; None where one of them reads a name of a sheet of that workbook, which the model does not follow."""
    order = [sheet.casefold() for sheet in linked.sheets]
    areas: dict[int, list[Area]] = {}

    def add(sheets: tuple[str, str], rest: str) -> bool:
        first, last = (sheet.casefold() for sheet in sheets)
        if first not in order or last not in order:
            return True
        try:
            found = parse_area(rest)
        except ValueError:
            return False
        low, high = sorted((order.index(first), order.index(last)))
        for index in range(low, high + 1):
            areas.setdefault(index, []).append(found)
        return True

    for formula, _ in _formulas(book):
        for reference in _references(formula):
            if reference.number != linked.number:
                continue
            if reference.first:
                if not add((reference.first, reference.last), reference.rest):
                    return None
                continue
            refers = next((text for key, text in linked.names.items() if key.casefold() == reference.rest.casefold()),
                          "")
            for piece in re.findall(r"((?:'(?:[^']|'')*'|[^'!=(),+\-*/&^<>]+)![$A-Z0-9:]+)", refers):
                sheet, area = split_sheet(piece)
                if not add((sheet, sheet), area):
                    return None
    return areas


def _drop_parts(book: Workbook, package: OpcFile, parts: list[str]) -> None:
    """The broken links' parts taken out of the package: each part and its relationships, its content type, the
    workbook's relationship to it and its externalReference."""
    if not parts:
        return
    from pyopenvba.apps.excel._charts_file import rels_part

    rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    types = package.read("[Content_Types].xml").decode("utf-8")
    for part in parts:
        if not part:
            continue
        target = part.removeprefix("xl/")
        found = re.search(r'<Relationship\b[^>]*\bTarget="/?(?:xl/)?' + re.escape(target) + r'"[^>]*/>', rels)
        if found is not None:
            identity = re.search(r'\bId="([^"]+)"', found.group(0))
            rels = rels[:found.start()] + rels[found.end():]
            if identity is not None:
                workbook = re.sub(r'<externalReference\b[^>]*\br:id="' + re.escape(identity.group(1)) + r'"[^>]*/>',
                                  "", workbook, count=1)
        types = re.sub(r'<Override\b[^>]*\bPartName="/' + re.escape(part) + r'"[^>]*/>', "", types, count=1)
        for name in (part, rels_part(part)):
            if package.has(name):
                package.remove(name)
    workbook = workbook.replace("<externalReferences></externalReferences>", "")
    package.write("xl/_rels/workbook.xml.rels", rels.encode("utf-8"))
    package.write("xl/workbook.xml", workbook.encode("utf-8"))
    package.write("[Content_Types].xml", types.encode("utf-8"))


def _number_parts(book: Workbook, package: OpcFile) -> None:
    """Each link's part named for its number, externalLink1.xml first, as Excel names them on every save."""
    from pyopenvba.apps.excel._charts_file import rels_part

    renames = {linked.part: f"xl/externalLinks/externalLink{linked.number}.xml" for linked in _links(book)
               if linked.part and linked.part != f"xl/externalLinks/externalLink{linked.number}.xml"}
    if not renames:
        return
    contents = {old: (package.read(old), package.read(rels_part(old)) if package.has(rels_part(old)) else None)
                for old in renames}
    for old in renames:
        package.remove(old)
        if package.has(rels_part(old)):
            package.remove(rels_part(old))
    rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    types = package.read("[Content_Types].xml").decode("utf-8")
    temporary = {old: f"\x00{index}\x00" for index, old in enumerate(renames)}
    for old, mark in temporary.items():
        rels = rels.replace(f'Target="{old.removeprefix("xl/")}"', f'Target="{mark}"').replace(
            f'Target="/{old}"', f'Target="{mark}"')
        types = types.replace(f'PartName="/{old}"', f'PartName="{mark}"')
    for old, new in renames.items():
        mark = temporary[old]
        rels = rels.replace(f'Target="{mark}"', f'Target="{new.removeprefix("xl/")}"')
        types = types.replace(f'PartName="{mark}"', f'PartName="/{new}"')
        data, relationships = contents[old]
        package.write(new, data)
        if relationships is not None:
            package.write(rels_part(new), relationships)
    package.write("xl/_rels/workbook.xml.rels", rels.encode("utf-8"))
    package.write("[Content_Types].xml", types.encode("utf-8"))
    book.linked_books = {number: replace(linked, part=renames.get(linked.part, linked.part))
                         for number, linked in book.linked_books.items()}
    changes = book.link_changes
    changes.values = {renames.get(part, part): text for part, text in changes.values.items()}
