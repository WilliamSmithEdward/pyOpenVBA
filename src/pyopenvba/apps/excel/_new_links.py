"""A link a macro's formula makes by naming a workbook the file has no link to, and the part a save writes for it.

A formula may name the workbook by its file, [Alpha.xlsx]Sheet1!A1, or by its path, 'C:\\dir\\[Alpha.xlsx]Sheet1'!A1,
and Excel makes a link, numbered after the links there are, in the order the formulas came. An open workbook of that
name is read as it stands, and naming a sheet it does not have is error 1004. A closed one is read from its file: by
the path given, or for a name alone, in the current folder. A file without the sheet has its first sheet read in its
place, and the formula names that sheet. A file not there leaves the link with the name it was given and the formula
at #REF!. A workbook never saved is linked by its name. The values stay once the workbook closes, and a link no formula
reads any more goes. Measured by scripts/measure_book_refs.py (tests/fixtures/book_refs.json).

A file on disk is read past the interpreter's gate, as UpdateLink and ChangeLink read one, so the sandbox refuses it
before anything is looked for (see _link_edits.loaded).
"""
from __future__ import annotations

import os
from dataclasses import replace
from html import escape
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._files import on_disk
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._external_links import LinkedBook
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.powerquery._opc import OpcFile

_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_LINK_PATH = _R + "/externalLinkPath"
_PATH_MISSING = "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing"
_LINK = _R + "/externalLink"
_LINK_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"
_HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<externalLink '
         'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
         'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14 xxl21" '
         'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
         'xmlns:xxl21="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021">')


def _open_source(book: Workbook, given: str, file: str) -> Workbook | None:
    """The open workbook a formula names: by its name, and in the folder given if one is."""
    folder = given.replace("/", "\\").rstrip("\\").casefold()
    return next((other for other in book.application.workbooks_.books
                 if other is not book and other.name.casefold() == file.casefold()
                 and (not folder or other.path.replace("/", "\\").rstrip("\\").casefold() == folder)), None)


def _current_folder(book: Workbook) -> str:
    interpreter = book.application.interpreter
    return interpreter.files.folder if interpreter is not None else os.getcwd()


def made(book: Workbook, given: str, file: str, sheet: str) -> str:
    """The new link a formula naming ``file``'s ``sheet`` makes, in the folder ``given`` or with none, as the formula
    then names it, [2]Sheet1; ``sheet`` empty for one of the workbook's own names, [2]."""
    from pyopenvba.apps.excel._external_links import LinkedBook
    from pyopenvba.apps.excel._link_edits import contents, loaded

    number = max(book.linked_books, default=0) + 1
    source = _open_source(book, given, file)
    if source is not None:
        found: tuple[Workbook, bool] | None = (source, True)
        target, pathless = str(source.FullName()), not source.path
    else:
        # Nothing on disk is looked at before the gate in loaded: off Windows on_disk reads folders for their case.
        target = os.path.join(given.rstrip("\\/") if given else _current_folder(book), file)
        found = loaded(book, target)
        if found is not None and not found[1]:
            target = on_disk(target)
        pathless = found is None and not given
    if found is None:
        if not sheet:
            raise VBAUnsupportedError("a name in a workbook whose file is not there is not implemented")
        # Read as #REF!: the link lists the sheet, as one its workbook does not have.
        book.linked_books[number] = LinkedBook(number, file if pathless else target, (), (), {}, extra_sheets=(sheet,),
                                               made=True, pathless=pathless)
        return f"[{number}]{sheet}"
    workbook, live = found
    sheets, cells = contents(workbook, live)
    if sheet:
        own = next((one for one in sheets if one.casefold() == sheet.casefold()), None)
        if own is None and live:
            raise error(1004, "Application-defined or object-defined error")
        sheet = own if own is not None else sheets[0]
    names = {entry.name: entry.refers_to for entry in workbook.names_.entries if entry.sheet is None}
    book.linked_books[number] = LinkedBook(number, target, sheets, cells, names, made=True, pathless=pathless)
    return f"[{number}]{sheet}"


def drop_unread(book: Workbook) -> None:
    """Take out each link a formula made that no formula or name reads any more, the later ones moving up."""
    from pyopenvba.apps.excel._workbook_links import formulas_of, references_in, renumber

    held = {reference.number for formula, _ in formulas_of(book) for reference in references_in(formula)}
    for number in sorted((number for number, linked in book.linked_books.items() if linked.made and number not in held),
                         reverse=True):
        part = book.linked_books[number].part
        if part:
            book.link_changes.broken.append(part)
        renumber(book, number)


# -- saving


def _relationships(linked: LinkedBook, here: str) -> tuple[str, str]:
    """A new link's relationships and alternateUrls, as Excel writes them for its workbook's place against ``here``,
    the folder the workbook is saved in: a workbook below it by its relative path, one elsewhere by its path from the
    root, each with its whole path beside it; one not found by the name it was given."""
    def one(identity: int, target: str, kind: str = _LINK_PATH) -> str:
        return f'<Relationship Id="rId{identity}" Type="{kind}" Target="{escape(target)}" TargetMode="External"/>'

    if linked.pathless:
        return one(1, linked.target, _PATH_MISSING), ""
    try:
        relative = os.path.relpath(linked.target, here).replace(os.sep, "/")
    except ValueError:
        raise VBAUnsupportedError("saving a link to a workbook on another drive is not implemented") from None
    whole = "file:///" + linked.target.lstrip("/")
    if not linked.sheets:
        if "/" in relative:
            raise VBAUnsupportedError("saving a link to a missing workbook in another folder is not implemented")
        return one(1, relative), ""
    if not relative.startswith("../"):
        return one(2, whole) + one(1, relative), '<xxl21:alternateUrls><xxl21:absoluteUrl r:id="rId2"/></xxl21:alternateUrls>'
    rooted = os.path.splitdrive(linked.target)[1].replace("\\", "/")
    return (one(3, relative) + one(2, whole) + one(1, rooted),
            '<xxl21:alternateUrls><xxl21:absoluteUrl r:id="rId2"/><xxl21:relativeUrl r:id="rId3"/></xxl21:alternateUrls>')


def _defined_names(book: Workbook, linked: LinkedBook) -> str:
    """The linked workbook's names a formula reads, [1]!Rate, each with its sheet in apostrophes, as Excel writes them."""
    import re

    from pyopenvba._a1 import split_sheet
    from pyopenvba.apps.excel._workbook_links import formulas_of

    read = {found.casefold() for formula, _ in formulas_of(book)
            for found in re.findall(rf"\[{linked.number}\]!([A-Za-z_\\][A-Za-z0-9_.\\]*)", formula)}
    entries: list[str] = []
    for name, refers in linked.names.items():
        if name.casefold() not in read:
            continue
        sheet, rest = split_sheet(refers.removeprefix("="))
        if not sheet or "!" in rest or "(" in rest:
            raise VBAUnsupportedError("saving a new link whose workbook has a name that is not one sheet's cells is "
                                      "not implemented")
        written = "='" + sheet.replace("'", "''") + "'!" + rest
        entries.append(f'<definedName name="{escape(name)}" refersTo="{escape(written, quote=False).replace(chr(34), "&quot;")}"/>')
    return f"<definedNames>{''.join(entries)}</definedNames>" if entries else ""


def write_made(book: Workbook, package: OpcFile, here: str) -> None:
    """Each link a formula made that no save has written yet, as its own part, listed in the workbook's."""
    import re

    from pyopenvba.apps.excel._charts_file import rels_part
    from pyopenvba.apps.excel._link_edits import data_set, source_is_open
    from pyopenvba.apps.excel._workbook_links import links_of, reads_book_names, with_book_sheet

    new = [linked for linked in links_of(book) if linked.made and not linked.part]
    if not new:
        return
    if not book.save_link_values:
        raise VBAUnsupportedError("saving a new link with SaveLinkValues False is not implemented")
    rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    types = package.read("[Content_Types].xml").decode("utf-8")
    taken = {int(found) for found in re.findall(r'\bId="rId([0-9]+)"', rels)}
    for linked in new:
        part = f"xl/externalLinks/externalLink{linked.number}.xml"
        if package.has(part):
            raise VBAUnsupportedError("a new link numbered where a link's part is already is not implemented")
        relationships, alternate = _relationships(linked, here)
        sheets = (*linked.sheets, *linked.extra_sheets)
        names = "".join(f'<sheetName val="{escape(sheet)}"/>' for sheet in sheets)
        package.write(rels_part(part), ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships '
                                        'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                        + relationships + "</Relationships>").encode("utf-8"))
        text = (f'{_HEAD}<externalBook xmlns:r="{_R}" r:id="rId1">{alternate}<sheetNames>{names}'
                f'</sheetNames>{_defined_names(book, linked)}{data_set(book, linked)}</externalBook></externalLink>')
        if reads_book_names(book, linked.number) and not source_is_open(book, linked):
            # As Excel saves a link to a closed workbook whose names a formula reads (case 16, not 6, of the fixture).
            text = with_book_sheet(linked, text)
        package.write(part, text.encode("utf-8"))
        identity = f"rId{max(taken, default=0) + 1}"
        taken.add(max(taken, default=0) + 1)
        rels = rels.replace("</Relationships>", f'<Relationship Id="{identity}" Type="{_LINK}" '
                                                f'Target="externalLinks/externalLink{linked.number}.xml"/></Relationships>')
        reference = f'<externalReference r:id="{identity}"/>'
        if "</externalReferences>" in workbook:
            workbook = workbook.replace("</externalReferences>", reference + "</externalReferences>", 1)
        else:
            workbook = workbook.replace("</sheets>", f"</sheets><externalReferences>{reference}</externalReferences>", 1)
        types = types.replace("</Types>", f'<Override PartName="/{part}" ContentType="{_LINK_TYPE}"/></Types>')
        book.linked_books[linked.number] = replace(linked, part=part)
    package.write("xl/_rels/workbook.xml.rels", rels.encode("utf-8"))
    package.write("xl/workbook.xml", workbook.encode("utf-8"))
    package.write("[Content_Types].xml", types.encode("utf-8"))
