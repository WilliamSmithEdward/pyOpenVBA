"""Links to other workbooks, and the copy of their cells Excel keeps in each.

A formula reads another workbook as ``[1]Sheet1!A1``, the number being the
link's place among the workbook part's ``<externalReferences>``, and names
one of its names as ``[1]!Rate``. Each link is a part of its own, whose
``<externalBook>`` lists the linked workbook's sheets and names and keeps the
value of every cell a formula read there, as it stood when the workbook was
last saved with the link up to date. Excel calculates with that copy while
the linked workbook is closed, and Range.Formula shows the link as the path
of the workbook it points at, ``'C:\\dir\\[Source.xlsx]Sheet1'!A1``
(tests/fixtures/external_links.json). Read as pyOfficeEditor reads it (its
``_externals.py``, commit 691adba).
"""
from __future__ import annotations

import ntpath
import os
import posixpath
import re
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING

from pyopenvba.formula._values import ERRORS, ExcelError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.powerquery._opc import OpcFile

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_RELATIONSHIPS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PACKAGE = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_LINK_PATH = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
_PATH_MISSING = "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing"
_CELL = re.compile(r"([A-Za-z]{1,3})([0-9]{1,7})")


@dataclass(frozen=True)
class LinkedBook:
    """A linked workbook as its link keeps it: where it is, its sheets in order, each sheet's cells by row and column,
    and its workbook-wide names with what each refers to, written as in that workbook."""

    number: int
    #: The workbook's path as the link records it, relative to this workbook's folder unless it is absolute.
    target: str
    sheets: tuple[str, ...]
    cells: tuple[dict[tuple[int, int], object], ...]
    names: dict[str, str]
    #: The package part the link was read from.
    part: str = ""
    #: Whether the link keeps the linked cells' values: a workbook saved with SaveLinkValues False keeps none.
    kept: bool = True
    #: The sheets a save lists after the workbook's own once ChangeLink has pointed the link at it (see _link_edits).
    extra_sheets: tuple[str, ...] = ()
    #: Made this session by a formula naming a workbook the file had no link to (see _new_links).
    made: bool = False
    #: Known by its file's name alone: a workbook never saved, or one a formula named whose file was not found.
    pathless: bool = False

    def folder(self, here: str) -> str:
        """The folder the linked workbook is in, ``here`` being this workbook's own."""
        if self.pathless:
            return ""
        # A POSIX path by its spelling: this workbook's folder from the root, or, for a workbook never saved, the system's own. A drive-less
        # target under a Windows folder takes that folder's drive on any system.
        posix = here.startswith("/") if here else os.name != "nt"
        if posix and self.target.startswith("/"):
            return posixpath.dirname(self.target)
        target = self.target.removeprefix("file:///").replace("/", "\\")
        head = target.rpartition("\\")[0]
        if re.match(r"[A-Za-z]:\\|\\\\", target) or not here:
            return head
        if here.startswith("/"):
            # A relative folder joins a POSIX folder as one, not turned into \tmp\x\Sub.
            return posixpath.normpath(posixpath.join(here, head.replace("\\", "/"))) if head else here.rstrip("/")
        return ntpath.normpath(ntpath.join(here, head)) if head else here.rstrip("\\")

    def path(self, here: str) -> str:
        """The linked workbook's full name, as LinkSources gives it."""
        folder = self.folder(here)
        return f"{folder}\\{self.file}" if folder else self.file

    @property
    def file(self) -> str:
        """The linked workbook's file name."""
        return self.target.replace("/", "\\").rpartition("\\")[2]


def read(book: Workbook, package: OpcFile, workbook: XmlElement) -> None:
    """The links ``workbook``, the parsed workbook part, lists, numbered as its formulas number them. A link that is
    not to a workbook, such as DDE or OLE, keeps its place and is left out."""
    from pyopenvba.apps.excel._charts_file import rels_part, xml_tree

    book.linked_books = {}
    book.foreign_links = False
    references = workbook.find(_MAIN + "externalReferences")
    if references is None:
        return
    relationships = _relationships(package, "xl/workbook.xml")
    for number, reference in enumerate(references.findall(_MAIN + "externalReference"), start=1):
        part = relationships.get(reference.get(_RELATIONSHIPS + "id", ""), ("", ""))[1]
        if not part or not package.has(part):
            book.foreign_links = True
            continue
        linked = xml_tree(package.read(part)).find(_MAIN + "externalBook")
        if linked is None:
            book.foreign_links = True
            continue
        paths = _relationships(package, part) if package.has(rels_part(part)) else {}
        kind, target = paths.get(linked.get(_RELATIONSHIPS + "id", ""), ("", ""))
        found = _linked(number, target if kind in (_LINK_PATH, _PATH_MISSING) else "", linked, part)
        # A workbook Excel knew by its name alone: one never saved, or not found (tests/fixtures/book_refs.json).
        book.linked_books[number] = replace(found, pathless=True) if kind == _PATH_MISSING else found


def _relationships(package: OpcFile, part: str) -> dict[str, tuple[str, str]]:
    """Each relationship of a part by its id: its type, and the part it names, or for an external one its target."""
    from pyopenvba.apps.excel._charts_file import rels_part, xml_tree

    found: dict[str, tuple[str, str]] = {}
    for entry in xml_tree(package.read(rels_part(part))).findall(_PACKAGE + "Relationship"):
        target = entry.get("Target", "")
        if entry.get("TargetMode") != "External":
            target = target[1:] if target.startswith("/") else posixpath.normpath(
                posixpath.join(posixpath.dirname(part), target))
        found[entry.get("Id", "")] = (entry.get("Type", ""), target)
    return found


def _children(parent: XmlElement, container: str, name: str) -> list[XmlElement]:
    found = parent.find(_MAIN + container)
    return [] if found is None else found.findall(_MAIN + name)


def _linked(number: int, target: str, linked: XmlElement, part: str) -> LinkedBook:
    from pyopenvba._a1 import column_number
    from pyopenvba.apps.excel._xstring import decode

    sheets = tuple(decode(entry.get("val", "")) for entry in _children(linked, "sheetNames", "sheetName"))
    cells: list[dict[tuple[int, int], object]] = [{} for _ in sheets]
    for sheet in _children(linked, "sheetDataSet", "sheetData"):
        index = sheet.get("sheetId", "")
        if not index.isdigit() or not int(index) < len(cells):
            continue
        for row in sheet.findall(_MAIN + "row"):
            for cell in row.findall(_MAIN + "cell"):
                value = _value(cell)
                found = _CELL.fullmatch(cell.get("r", ""))
                if value is not None and found is not None:
                    cells[int(index)][(int(found.group(2)), column_number(found.group(1)))] = value
    names = {decode(entry.get("name", "")): decode(entry.get("refersTo", ""))
             for entry in _children(linked, "definedNames", "definedName") if entry.get("sheetId") is None}
    kept = linked.find(_MAIN + "sheetDataSet") is not None
    return LinkedBook(number, target, sheets, tuple(cells), names, part, kept)


def _value(cell: XmlElement) -> object | None:
    """A kept cell's value: a number unless its type says text (str), a logical (b) or an error (e)."""
    from pyopenvba.apps.excel._xstring import decode

    found = cell.find(_MAIN + "v")
    if found is None:
        return None
    raw = decode(found.text or "")
    kind = cell.get("t", "n")
    if kind == "str":
        return raw
    if kind == "b":
        return raw in ("1", "true")
    if kind == "e":
        return ERRORS.get(raw.upper(), ExcelError(raw.upper()))
    try:
        return float(raw)
    except ValueError:
        return None


#: A reference or name to a linked workbook, by its number: [1]Sheet1!, [1]Sheet1:Other!, [1]!.
_PREFIX = re.compile(r"\[([0-9]+)\]((?:[^'!\[\]]|'')*)!")


def linked_sheet(book: Workbook, name: str) -> str:
    """Another workbook's sheet a macro's formula names, by its path, C:\\dir\\[Source.xlsx]Sheet1, or its file's name,
    [Source.xlsx]Sheet1, as the file names it by its link, [1]Sheet1; a workbook alone, C:\\dir\\Source.xlsx for one of
    its names, is [1]. Excel takes either for a workbook it has a link to (tests/fixtures/external_links.json); one it
    has none to makes a new link (see _new_links)."""
    from pyopenvba.apps.excel._new_links import made

    if name.startswith("[") and name[1:].partition("]")[0].isdigit():
        return name
    given, bracket, rest = name.partition("[")
    if bracket:
        file, _, sheet = rest.partition("]")
    else:
        given, file = ntpath.split(name)
        sheet = ""
    # Off Windows the folder a macro names has / in it, and so does the one a link's relative target comes to.
    folder = given.replace("/", "\\").rstrip("\\").casefold()
    for linked in book.linked_books.values():
        if linked.file.casefold() == file.casefold() \
                and (not folder or linked.folder(book.path).replace("/", "\\").casefold() == folder):
            return f"[{linked.number}]{sheet}"
    return made(book, given, file, sheet)


def shown(book: Workbook, formula: str) -> str:
    """A formula's references to linked workbooks as Range.Formula shows them: each [1]Sheet1! as the path of the
    workbook the link points at, 'C:\\dir\\[Source.xlsx]Sheet1'!, and each [1]! of a name as 'C:\\dir\\Source.xlsx'!.
    A link the workbook does not list is left as it is."""
    from pyopenvba.apps.excel._link_edits import source_is_open
    from pyopenvba._a1 import quote_sheet
    from pyopenvba.formula import _parse as P

    if "[0]!" in formula:
        # This workbook, by what a deleted sheet leaves of its name: a file holds [0]!Name, and Excel shows the workbook's
        # name as it is now (tests/fixtures/deleted_local_names.json).
        formula = formula.replace("[0]!", quote_sheet(book.name) + "!")
    if "[" not in formula or not getattr(book, "linked_books", None):
        return formula
    try:
        tokens = P.tokenize(formula[1:], spaces=True) if formula.startswith("=") else None
    except P.FormulaError:
        return formula
    if tokens is None:
        return formula
    out: list[str] = []
    for token in tokens:
        text = token.text
        found = _PREFIX.match(text) if token.kind in ("ref", "name") else None
        linked = book.linked_books.get(int(found.group(1))) if found is not None else None
        if found is not None and linked is not None:
            sheet = found.group(2)
            if source_is_open(book, linked) or linked.pathless:
                # A workbook known by its name alone is shown so even once closed (tests/fixtures/book_refs.json).
                place = f"[{linked.file}]{sheet}" if sheet else linked.file
                quoted = re.search(r"[^A-Za-z0-9_.\[\]]", place) is not None
                text = ("'" + place.replace("'", "''") + "'" if quoted else place) + "!" + text[found.end():]
            else:
                folder = linked.folder(book.path)
                place = (folder + "\\" if folder else "") + (f"[{linked.file}]{sheet}" if sheet else linked.file)
                text = "'" + place.replace("'", "''") + "'!" + text[found.end():]
        out.append(text)
    return "=" + "".join(out)
