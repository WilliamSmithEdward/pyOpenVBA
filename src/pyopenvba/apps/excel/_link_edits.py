"""Workbook.UpdateLink, ChangeLink, OpenLinks and SetLinkOnData: what a macro does to a link's workbook.

UpdateLink reads a closed workbook's cells into the link. ChangeLink points a link at another workbook and reads it;
one the file is not there for leaves its formulas at #REF!, and a workbook without the sheets read has its first sheet
read in their place. The links are then numbered in the order of their full names, and a link pointed at a workbook
another already reads joins it. OpenLinks opens files, which need not be the links', on a workbook that has links.
SetLinkOnData is for DDE and OLE links and does nothing here. While a linked workbook is open, formulas name it by its
file alone and read its cells as they stand; when it closes they keep what it last held. Measured in live Excel by
scripts/measure_link_edits.py (tests/fixtures/link_edits.json).
"""
from __future__ import annotations

import ntpath
import posixpath
import re
from collections.abc import Callable
from dataclasses import replace
from html import escape
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._hooks import link_gate
from pyopenvba.apps.excel._workbook_links import (
    found_link,
    link_kind,
    links_of,
    links_read,
    read_areas,
    read_sheets,
    references_in,
    refuse_elsewhere,
    renumber,
)
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._files import on_disk
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAArray, error, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._external_links import LinkedBook
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.powerquery._opc import OpcFile

_EXCEL_LINKS, _OLE_LINKS = 1, 2
_LINK_PATH = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _failed(member: str) -> Exception:
    return error(1004, f"{member} method of Workbook class failed")


def _same_path(path: str) -> str:
    """A path to compare: no case, and either separator."""
    return path.replace("/", "\\").casefold()


def _open_book(book: Workbook, path: str) -> Workbook | None:
    """The open workbook at ``path``, the path a link gives, whose last separator is a backslash on any system."""
    wanted = _same_path(path)
    return next((other for other in book.application.workbooks_.books
                 if other is not book and _same_path(str(other.FullName())) == wanted), None)


def source_is_open(book: Workbook, linked: LinkedBook) -> bool:
    return _open_book(book, linked.path(book.path)) is not None


def loaded(book: Workbook, path: str) -> tuple[Workbook, bool] | None:
    """The workbook at ``path`` and whether it is open, or None where it is neither open nor on disk. One on disk is
    read past the interpreter's gate, which the sandbox shuts."""
    from pyopenvba.apps.excel._io import load_workbook

    source = _open_book(book, path)
    if source is not None:
        return source, True
    token = link_gate(book.application, "Reading a linked workbook", path)
    file = Path(on_disk(path))
    if not file.is_file():
        return None
    return load_workbook(book.application, file, _passed=token), False


def _read(book: Workbook, path: str) -> tuple[tuple[str, ...], tuple[dict[tuple[int, int], object], ...]] | None:
    """The sheets and cells of the workbook at ``path``, open or on disk, or None where it is neither."""
    found = loaded(book, path)
    return None if found is None else contents(*found)


def contents(source: Workbook, live: bool) -> tuple[tuple[str, ...], tuple[dict[tuple[int, int], object], ...]]:
    """A workbook's sheets and the values of their cells: as they stand where it is open, as its file kept them where
    not."""
    from pyopenvba.apps.excel._engine_book import model_value, scalar
    from pyopenvba.formula._calc.values import Empty

    names: list[str] = []
    cells: list[dict[tuple[int, int], object]] = []
    for sheet in source.sheets_:
        names.append(sheet.name)
        kept: dict[tuple[int, int], object] = {}
        for (row, column), cell in sheet.cells_.items():
            value = source.calculator.value_of(sheet.name, row, column) if live and cell.formula else cell.value
            held = scalar(value)
            if not isinstance(held, Empty):
                kept[(row, column)] = model_value(held)
        cells.append(kept)
    return tuple(names), tuple(cells)


def _refreshed(book: Workbook, linked: LinkedBook) -> bool:
    """The link's copy read again from its workbook; whether the workbook was there."""
    found = _read(book, linked.path(book.path))
    if found is None:
        return False
    book.linked_books[linked.number] = replace(linked, sheets=found[0], cells=found[1], kept=True)
    return True


def _readers(book: Workbook, numbers: set[int]) -> list[tuple[str, int, int]]:
    """The cells whose formulas read any of the links ``numbers``, themselves or through a name."""
    return [(sheet.name, row, column) for sheet in book.sheets_ for (row, column), cell in sheet.cells_.items()
            if cell.formula and links_read(book, cell.formula, sheet) & numbers]


def _wanted(value: object) -> list[object]:
    """A link's Name or the array of them."""
    if isinstance(value, VBAArray):
        return list(value.items)
    return [value]


def update_link(book: Workbook, name: object, kind: object) -> object:
    """Workbook.UpdateLink. With neither argument it fails. An empty name, or none with a Type, is every link; a
    workbook open or gone, and a name that is no link, fail, and one broken already is left as it is."""
    if name is MISSING and kind is MISSING:
        raise _failed("UpdateLink")
    which = link_kind(kind)
    if which not in (_EXCEL_LINKS, _OLE_LINKS):
        raise _failed("UpdateLink")
    every = name is MISSING or name is EMPTY or name == ""
    if which == _OLE_LINKS:
        if every:
            return True
        raise _failed("UpdateLink")
    link_gate(book.application, "Workbook.UpdateLink")
    if every:
        chosen = [linked for linked in links_of(book) if not source_is_open(book, linked)]
    else:
        chosen: list[LinkedBook] = []
        for one in _wanted(name):
            if not isinstance(one, str):
                raise _failed("UpdateLink")
            linked, broken = found_link(book, one)
            if linked is None or (not broken and source_is_open(book, linked)):
                raise _failed("UpdateLink")
            if not broken:
                chosen.append(linked)
    readers = _readers(book, {linked.number for linked in chosen})
    for linked in chosen:
        if not _refreshed(book, linked):
            raise _failed("UpdateLink")
        book.link_changes.reread.add(linked.part)
    if chosen:
        book.calculator.reread(readers)
        book.saved = False
    return True


def set_link_on_data(name: object) -> object:
    """Workbook.SetLinkOnData: for DDE and OLE links, so nothing happens to a workbook's own."""
    if name is MISSING:
        raise error(449, "Argument not optional")
    return True


def open_links(book: Workbook, name: object, read_only: object, kind: object) -> object:
    """Workbook.OpenLinks: on a workbook with links, the files named, which need not be its links' files."""
    if name is MISSING:
        raise error(449, "Argument not optional")
    book.application.outside("Workbook.OpenLinks")
    if link_kind(kind) != _EXCEL_LINKS or not links_of(book):
        raise _failed("OpenLinks")
    for one in _wanted(name):
        text = to_text(one)
        given = text if isinstance(one, str) else f"{text}.xlsx"
        if not Path(on_disk(given)).is_file():
            raise error(1004, f"Sorry, we couldn't find {given}. Is it possible it was moved, renamed or deleted?")
        book.application.workbooks_.Open(given, ReadOnly=read_only if read_only is not MISSING else False)
    return True


def change_link(book: Workbook, name: object, new_name: object, kind: object) -> object:
    """Workbook.ChangeLink; see the module's docstring. A name that is no link, or an empty new one, fails and brings
    the workbook to the front; pointing a link where it points already, or at this workbook, does nothing."""
    if name is MISSING or new_name is MISSING:
        raise error(449, "Argument not optional")
    link_gate(book.application, "Workbook.ChangeLink")
    if link_kind(kind) != _EXCEL_LINKS:
        raise _failed("ChangeLink")
    target = to_text(new_name)
    linked, broken = found_link(book, to_text(name)) if isinstance(name, str) else (None, False)
    if linked is None or not target:
        book.application.activate_book(book)
        raise _failed("ChangeLink")
    if broken:
        return True
    # Off Windows a full path starts with /, which ntpath does not call absolute and normpath would turn round.
    posix = target.startswith("/")
    if not posix and not ntpath.isabs(target):
        raise VBAUnsupportedError("ChangeLink to a workbook named without its folder is not implemented")
    target = posixpath.normpath(target.replace("\\", "/")) if posix else ntpath.normpath(target)
    own = _same_path(str(book.FullName()))
    if _same_path(target) in (_same_path(linked.path(book.path)), own):
        return True
    if book.foreign_links:
        raise VBAUnsupportedError("ChangeLink in a workbook with OLE or DDE links is not implemented")
    refuse_elsewhere(book)
    _refuse_book_names(book, linked.number)
    sheets_read = read_sheets(book, linked)
    found = _read(book, target)
    names, cells = found if found is not None else ((), ())
    first = names[0] if names else None
    lost = {sheet.casefold() for sheet in sheets_read if first is not None
            and sheet.casefold() not in {one.casefold() for one in names}}
    joined = next((other for other in links_of(book) if other is not linked
                   and _same_path(other.path(book.path)) == _same_path(target)), None)
    number = joined.number if joined is not None else linked.number
    readers = _readers(book, {linked.number})
    gone = linked.number

    def repointed(at: int, top: str, bottom: str) -> tuple[int, str, str]:
        if at != gone:
            return at, top, bottom
        return number, first if first is not None and top.casefold() in lost else top, \
            first if first is not None and bottom.casefold() in lost else bottom

    _rewrite(book, repointed)
    stem = ntpath.basename(target).rpartition(".")[0]
    extra = tuple(dict.fromkeys(sheet for sheet in (stem, *sorted(sheets_read, key=linked.sheets.index))
                                if sheet.casefold() not in {one.casefold() for one in names}))
    base = joined if joined is not None else linked
    book.linked_books[number] = replace(base, target=_target_for(book, target), sheets=names, cells=cells, kept=True,
                                         names={}, extra_sheets=extra)
    changes = book.link_changes
    if joined is not None:
        changes.broken.append(linked.part)
        renumber(book, linked.number)
    _in_name_order(book)
    changes.retargeted.update(one.part for one in links_of(book))
    book.calculator.reread(readers)
    book.saved = False
    return True


def _target_for(book: Workbook, path: str) -> str:
    """A link's target, its file alone where it is in this workbook's folder and the whole path where not."""
    folder, file = ntpath.split(path)
    return file if _same_path(folder) == _same_path(book.path.rstrip("\\")) else path


def _refuse_book_names(book: Workbook, number: int) -> None:
    from pyopenvba.apps.excel._workbook_links import formulas_of

    for formula, _ in formulas_of(book):
        for reference in references_in(formula):
            if reference.number == number and (not reference.first or reference.first != reference.last):
                raise VBAUnsupportedError("ChangeLink of a link read through its workbook's names or across its "
                                          "sheets is not implemented")


def _rewrite(book: Workbook, change: Callable[[int, str, str], tuple[int, str, str]]) -> None:
    """Every reference to a linked workbook in the workbook's formulas and names as ``change`` gives it, from its
    number and its first and last sheets to theirs."""
    def moved(formula: str) -> str:
        body = formula[1:] if formula.startswith("=") else formula
        out: list[str] = []
        at = 0
        for reference in references_in(formula):
            number, first, last = change(reference.number, reference.first, reference.last)
            if (number, first, last) == (reference.number, reference.first, reference.last):
                continue
            sheets = first if first == last else f"{first}:{last}"
            quoted = re.search(r"[^A-Za-z0-9_.]", sheets) is not None
            text = f"[{number}]{sheets}"
            text = "'" + text.replace("'", "''") + "'" if quoted else text
            out += [body[at:reference.start], f"{text}!{reference.rest}"]
            at = reference.start + reference.length
        return formula if not out else ("=" if formula.startswith("=") else "") + "".join(out) + body[at:]

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


def _in_name_order(book: Workbook) -> None:
    """The links numbered in the order of their full names, each number keeping its place's part."""
    links = links_of(book)
    ordered = sorted(links, key=lambda linked: linked.path(book.path).casefold())
    if ordered == links:
        return
    places = {linked.number: index for index, linked in enumerate(ordered, start=1)}
    _rewrite(book, lambda number, first, last: (places.get(number, number), first, last))
    book.linked_books = {index: replace(linked, number=index, part=links[index - 1].part)
                         for index, linked in enumerate(ordered, start=1)}


# -- while a linked workbook is open


def opened(book: Workbook) -> None:
    """A workbook opened: the links other workbooks have to it read it."""
    for other in book.application.workbooks_.books:
        if other is book or not other.linked_books:
            continue
        for linked in links_of(other):
            if source_is_open(other, linked) and _open_book(other, linked.path(other.path)) is book:
                readers = _readers(other, {linked.number})
                _refreshed(other, linked)
                other.link_changes.reread.add(linked.part)
                other.calculator.reread(readers)


def source_changed(book: Workbook) -> None:
    """A cell of an open workbook changed: what other workbooks' links read of it needs another look."""
    books = book.application.workbooks_.books
    if len(books) < 2:
        return
    for other in books:
        if other is book or not other.linked_books:
            continue
        numbers = {linked.number for linked in other.linked_books.values()
                   if _open_book(other, linked.path(other.path)) is book}
        if numbers:
            other.link_changes.source_stale = True
            other.calculator.reread(_readers(other, numbers))


def settled(book: Workbook) -> None:
    """The links to open workbooks read again, if one changed since they last were."""
    changes = book.link_changes
    if changes.source_stale:
        changes.source_stale = False
        for linked in links_of(book):
            if source_is_open(book, linked):
                _refreshed(book, linked)
                changes.reread.add(linked.part)


def closing(book: Workbook) -> None:
    """A workbook about to close: the links to it keep what it holds now."""
    for other in book.application.workbooks_.books:
        if other is not book and other.linked_books:
            settled(other)


# -- saving


def _cell_xml(reference: str, value: object) -> str:
    from pyopenvba.apps.excel._styles import excel_number

    if isinstance(value, bool):
        return f'<cell r="{reference}" t="b"><v>{int(value)}</v></cell>'
    if isinstance(value, str):
        return f'<cell r="{reference}" t="str"><v>{escape(value, quote=False)}</v></cell>'
    if isinstance(value, ExcelError):
        return f'<cell r="{reference}" t="e"><v>{escape(value.name, quote=False)}</v></cell>'
    assert isinstance(value, float)
    return f'<cell r="{reference}"><v>{excel_number(value)}</v></cell>'


def data_set(book: Workbook, linked: LinkedBook) -> str:
    """A link's sheetDataSet: the cells read of each sheet of its workbook, row by row, and for a sheet it does not
    have, an empty sheetData marked refreshError."""
    from pyopenvba._a1 import column_letter

    areas = read_areas(book, linked)
    if areas is None:
        raise VBAUnsupportedError("saving a link read through a name of one of its workbook's sheets, after its "
                                  "workbook was read again, is not implemented")
    sheets: list[str] = []
    for index, cells in enumerate(linked.cells):
        rows: dict[int, list[tuple[int, object]]] = {}
        for (row, column), value in cells.items():
            if any(area.contains(row, column) for area in areas.get(index, [])):
                rows.setdefault(row, []).append((column, value))
        body = "".join(f'<row r="{row}">' + "".join(_cell_xml(f"{column_letter(column)}{row}", value)
                                                    for column, value in sorted(rows[row])) + "</row>"
                       for row in sorted(rows))
        sheets.append(f'<sheetData sheetId="{index}">{body}</sheetData>' if body else f'<sheetData sheetId="{index}"/>')
    sheets += [f'<sheetData sheetId="{index}" refreshError="1"/>'
               for index in range(len(linked.cells), len(linked.cells) + len(linked.extra_sheets))]
    return "<sheetDataSet>" + "".join(sheets) + "</sheetDataSet>"


def write_reread(book: Workbook, linked: LinkedBook, text: str) -> str:
    """A link's part with the values its workbook was last read for."""
    return re.sub(r"<sheetDataSet>.*?</sheetDataSet>|<sheetDataSet/>", lambda _: data_set(book, linked), text, count=1,
                  flags=re.DOTALL)


def write_retargeted(book: Workbook, package: OpcFile, linked: LinkedBook, text: str) -> str:
    """A link's part and relationships for the workbook it now points at, as Excel writes them: the sheets it has, a
    sheet named after it, and the sheets read that it does not have; and its absolute address only if the file is there."""
    from pyopenvba.apps.excel._charts_file import rels_part

    if "\\" in linked.target:
        raise VBAUnsupportedError("saving a link to a workbook in another folder is not implemented")
    path = linked.path(book.path)
    exists = bool(linked.cells)
    names = "".join(f'<sheetName val="{escape(sheet)}"/>' for sheet in (*linked.sheets, *linked.extra_sheets))
    alternate = '<xxl21:alternateUrls><xxl21:absoluteUrl r:id="rId2"/></xxl21:alternateUrls>' if exists else ""
    head = text[:text.index("<externalBook")]
    relationships = f'<Relationship Id="rId1" Type="{_LINK_PATH}" Target="{escape(linked.target)}" TargetMode="External"/>'
    if exists:
        relationships = f'<Relationship Id="rId2" Type="{_LINK_PATH}" Target="file:///{escape(path.lstrip("/"))}" TargetMode="External"/>' \
            + relationships
    package.write(rels_part(linked.part), ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships '
                                           'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                           + relationships + "</Relationships>").encode("utf-8"))
    return (f'{head}<externalBook xmlns:r="{_R}" r:id="rId1">{alternate}<sheetNames>{names}</sheetNames>'
            f'{data_set(book, linked)}</externalBook></externalLink>')

