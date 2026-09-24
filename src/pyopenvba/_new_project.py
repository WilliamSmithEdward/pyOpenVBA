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
_RELATIONSHIP = re.compile(r"<Relationship\b[^>]*?(?:/>|>\s*</Relationship>)")
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
    return [_attributes(element) for element in _RELATIONSHIP.findall(text)]


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
    rels = read(rels_name).decode("utf-8")
    numbers = [int(found.group(1)) for found in re.finditer(r'\bId="rId(\d+)"', rels)]
    relationship = (f'<Relationship Id="rId{max(numbers, default=0) + 1}" Type="{PROJECT_TYPE}" '
                    'Target="vbaProject.bin"/>')
    edits = {rels_name: rels.replace("</Relationships>", relationship + "</Relationships>", 1).encode("utf-8")}
    types = read(CONTENT_TYPES).decode("utf-8")
    defaults = list(re.finditer(r"<Default\b[^>]*/>", types))
    by_extension = {_attributes(found.group()).get("Extension", "").casefold(): found for found in defaults}
    root = re.search(r"<Types\b[^>]*>", types)
    assert root is not None
    if "bin" not in by_extension:
        later = [found for found in defaults if _attributes(found.group()).get("Extension", "").casefold() > "bin"]
        at = later[0].start() if later else (defaults[-1].end() if defaults else root.end())
        types = f'{types[:at]}<Default Extension="bin" ContentType="{PROJECT_CONTENT}"/>{types[at:]}'
    elif _attributes(by_extension["bin"].group()).get("ContentType") != PROJECT_CONTENT:
        # bin already names another type, printer settings say, so the project gets an Override of its own. Where
        # Excel puts one has not been measured: it goes first among the Overrides.
        at = defaults[-1].end()
        override = f'<Override PartName="/{folder}/vbaProject.bin" ContentType="{PROJECT_CONTENT}"/>'
        types = types[:at] + override + types[at:]
    edits[CONTENT_TYPES] = types.encode("utf-8")
    return edits
