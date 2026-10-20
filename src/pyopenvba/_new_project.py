"""A VBA project for a file that has none, as Excel, Word and PowerPoint make one.

Measured by scripts/measure_first_macro.py, which had each application
add a module to a file saved with no VBA, and add and remove one, and
save (tests/fixtures/first_macro/):

- Excel names a document module for the workbook and for each sheet in
  sheet order, one count across worksheets and chart sheets: Sheet1,
  Sheet2, Chart3, Sheet4. It writes the names into the workbook as
  codeName on workbookPr and on each sheet's sheetPr, a sheetPr coming
  first in a sheet that had none. It writes vbaProject.bin only once the
  project holds code: a project of document modules alone leaves the
  names and nothing more.
- Word makes ThisDocument, from the Normal template, and writes the
  project even when that is all it holds.
- PowerPoint makes no module of its own, and writes no project that
  holds none.

A project that is written is related from the main part's relationships
with a type of its own, and [Content_Types].xml gains a Default for the
bin extension, in extension order. Three things differ from what the
applications write. Office puts the relationship where its own ordering
puts it, and Word numbers every relationship again; here the
relationship is appended with the next free id, which Office accepts and
renumbers on its next save. Excel stamps fileVersion with a codeName
GUID that is a constant of its build, not of the file, so it is left for
Excel to write. And Word writes vbaData.xml once there is a macro, which
the library does not write for any module it adds.
"""

from __future__ import annotations

import posixpath
import re
from collections.abc import Callable, Sequence

from pyopenvba._opc_xml import RELS_NS, TYPES_NS, opc_part
from pyopenvba.exceptions import UnsupportedFormatError

PROJECT_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
PROJECT_CONTENT = "application/vnd.ms-office.vbaProject"
CONTENT_TYPES = "[Content_Types].xml"
_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
_CHARTSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
#: The VB_Base of each kind of Excel document module, as Excel writes it.
WORKBOOK_BASE = "0{00020819-0000-0000-C000-000000000046}"
_SHEET_BASES = {_WORKSHEET: ("Sheet", "0{00020820-0000-0000-C000-000000000046}"),
                _CHARTSHEET: ("Chart", "0{00020821-0000-0000-C000-000000000046}")}
#: Attributes that come before codeName in each element, in schema order; the rest come after.
_BEFORE_WORKBOOK_CODE_NAME = ("date1904", "showObjects", "showBorderUnselectedTables", "filterPrivacy",
                              "promptedSolutions", "showInkAnnotation", "backupFile", "saveExternalLinkValues",
                              "updateLinks")
_BEFORE_SHEET_CODE_NAME = ("syncHorizontal", "syncVertical", "syncRef", "transitionEvaluation", "transitionEntry",
                           "published")
_ATTRIBUTE = re.compile(r'([\w:]+)="([^"]*)"')
_START = re.compile(r"<(?!\?)(?!/)[^>]*>")

Reader = Callable[[str], bytes]


def document_header(name: str, base: str) -> str:
    """The attribute header Excel writes for a workbook or sheet document module."""
    return (f'Attribute VB_Name = "{name}"\r\nAttribute VB_Base = "{base}"\r\n'
            "Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
            "Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n"
            "Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n")


def _attributes(tag: str) -> dict[str, str]:
    return dict(_ATTRIBUTE.findall(tag))


def _with_attribute(tag: str, name: str, value: str, before: Sequence[str]) -> str:
    """``tag`` with ``name`` set: after the attributes schema order puts before it, else first."""
    present = list(_ATTRIBUTE.finditer(tag))
    earlier = [match for match in present if match.group(1) in before]
    if earlier:
        at = earlier[-1].end()
    elif present:
        at = present[0].start() - 1
    else:
        at = len(tag) - (2 if tag.endswith("/>") else 1)
    return f'{tag[:at]} {name}="{value}"{tag[at:]}'


def _relationships(text: str) -> list[dict[str, str]]:
    part = opc_part(text.encode("utf-8"), RELS_NS, "Relationships")
    return [attrs for name, attrs, _, _ in part.children if name == RELS_NS + "}Relationship"]


def _part(folder: str, target: str) -> str:
    return target[1:] if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))


def excel_sheets(names: Sequence[str], read: Reader) -> list[tuple[str, str, str]]:
    """Each sheet in the workbook's order: its part, its relationship type, and the code name it already has."""
    actual = {name.casefold(): name for name in names}
    workbook = read("xl/workbook.xml").decode("utf-8")
    targets = {found.get("Id", ""): found for found in _relationships(read("xl/_rels/workbook.xml.rels").decode("utf-8"))}
    sheets: list[tuple[str, str, str]] = []
    for element in re.findall(r"<sheet\b[^>]*/>", workbook):
        attributes = _attributes(element)
        relationship = targets.get(attributes.get("r:id", ""), {})
        kind = relationship.get("Type", "")
        if kind not in _SHEET_BASES:
            raise UnsupportedFormatError(
                f"Adding a VBA project to a workbook with a sheet of type {kind.rsplit('/', 1)[-1]!r} has not "
                "been measured.")
        part = actual.get(_part("xl", relationship.get("Target", "")).casefold(), "")
        text = read(part).decode("utf-8") if part else ""
        found = re.search(r"<sheetPr\b[^>]*>", text)
        sheets.append((part, kind, _attributes(found.group()).get("codeName", "") if found else ""))
    return sheets


def excel_code_names(names: Sequence[str], read: Reader) -> tuple[str, list[tuple[str, str, str]]]:
    """The workbook's code name, and each sheet's part, code name and document module base.

    A sheet keeps a code name it has; one without is named as Excel
    names it, by its place among the sheets.
    """
    workbook = read("xl/workbook.xml").decode("utf-8")
    found = re.search(r"<workbookPr\b[^>]*>", workbook)
    book_name = (_attributes(found.group()).get("codeName", "") if found else "") or "ThisWorkbook"
    sheets = excel_sheets(names, read)
    taken = {book_name.casefold()} | {code.casefold() for _, _, code in sheets if code}
    named: list[tuple[str, str, str]] = []
    for place, (part, kind, code) in enumerate(sheets, start=1):
        prefix, base = _SHEET_BASES[kind]
        if not code:
            code = f"{prefix}{place}"
            if code.casefold() in taken:
                raise UnsupportedFormatError(
                    f"Adding a VBA project where the code name {code!r} is already taken has not been measured.")
            taken.add(code.casefold())
        named.append((part, code, base))
    return book_name, named


def excel_code_name_edits(names: Sequence[str], read: Reader) -> dict[str, bytes]:
    """The workbook and sheet parts with the code names Excel writes, for those that lack one."""
    book_name, sheets = excel_code_names(names, read)
    edits: dict[str, bytes] = {}
    workbook = read("xl/workbook.xml").decode("utf-8")
    found = re.search(r"<workbookPr\b[^>]*>", workbook)
    if found is None:
        anchor = re.search(r"<fileSharing\b[^>]*/>|<fileVersion\b[^>]*/>|<workbook\b[^>]*>", workbook)
        assert anchor is not None
        workbook = f'{workbook[:anchor.end()]}<workbookPr codeName="{book_name}"/>{workbook[anchor.end():]}'
    elif "codeName" not in _attributes(found.group()):
        tag = _with_attribute(found.group(), "codeName", book_name, _BEFORE_WORKBOOK_CODE_NAME)
        workbook = workbook[:found.start()] + tag + workbook[found.end():]
    if workbook != read("xl/workbook.xml").decode("utf-8"):
        edits["xl/workbook.xml"] = workbook.encode("utf-8")
    for part, code, _ in sheets:
        text = read(part).decode("utf-8")
        found = re.search(r"<sheetPr\b[^>]*>", text)
        if found is None:
            root = next(match for match in _START.finditer(text))
            text = f'{text[:root.end()]}<sheetPr codeName="{code}"/>{text[root.end():]}'
        elif "codeName" not in _attributes(found.group()):
            tag = _with_attribute(found.group(), "codeName", code, _BEFORE_SHEET_CODE_NAME)
            text = text[:found.start()] + tag + text[found.end():]
        else:
            continue
        edits[part] = text.encode("utf-8")
    return edits


def with_project(names: Sequence[str], read: Reader, main_part: str) -> dict[str, bytes]:
    """The main part's relationships and [Content_Types].xml once they take the project beside ``main_part``."""
    folder, _, name = main_part.rpartition("/")
    rels_name = f"{folder}/_rels/{name}.rels"
    rels = opc_part(read(rels_name), RELS_NS, "Relationships")
    numbers = [int(match.group(1)) for name, attrs, _, _ in rels.children
               if name == RELS_NS + "}Relationship"
               and (match := re.fullmatch(r"rId(\d+)", attrs.get("Id", "")))]
    relationship = rels.child("Relationship", f'Id="rId{max(numbers, default=0) + 1}" '
                              f'Type="{PROJECT_TYPE}" Target="vbaProject.bin"')
    edits = {rels_name: rels.insert(rels.closing_start, relationship)}
    types = opc_part(read(CONTENT_TYPES), TYPES_NS, "Types")
    defaults = [(attrs, start, end) for name, attrs, start, end in types.children
                if name == TYPES_NS + "}Default"]
    by_extension = {attrs.get("Extension", "").casefold(): attrs for attrs, _, _ in defaults}
    if "bin" not in by_extension:
        later = [start for attrs, start, _ in defaults if attrs.get("Extension", "").casefold() > "bin"]
        at = later[0] if later else (defaults[-1][2] if defaults else types.opening_end)
        markup = types.child("Default", f'Extension="bin" ContentType="{PROJECT_CONTENT}"')
        edits[CONTENT_TYPES] = types.insert(at, markup)
    elif by_extension["bin"].get("ContentType") != PROJECT_CONTENT:
        # A bin Default for another part needs a project-specific Override.
        at = defaults[-1][2]
        markup = types.child("Override", f'PartName="/{folder}/vbaProject.bin" ContentType="{PROJECT_CONTENT}"')
        edits[CONTENT_TYPES] = types.insert(at, markup)
    else:
        edits[CONTENT_TYPES] = types.data
    return edits
