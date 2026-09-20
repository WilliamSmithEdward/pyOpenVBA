"""Reading a workbook into the model, and writing the model back out.

Loading reads the sheets, their cells, the defined names and the Power
Query queries.  Saving puts back only what changed: every part the model
does not describe, and every cell nobody touched, is carried over
verbatim, so a merge, a conditional format or a pivot table survives a
macro that wrote one cell.

That is the same rule the rest of pyOpenVBA follows.  Regenerating a
sheet from the model would be far easier and would quietly throw away
whatever the model has no field for.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, column_letter, column_number
from pyopenvba._xml import attributes as _attributes
from pyopenvba._xml import escape as _escape
from pyopenvba._xml import tag_attributes as _tag_attributes
from pyopenvba._xml import unescape as _unescape
from pyopenvba.exceptions import PyOpenVBAError
from pyopenvba.interpreter._values import EMPTY, VBADate, VBAInt, to_text
from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.powerquery._sheets import add_content_type, sheet_entries

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Cell, Workbook, Worksheet
    from pyopenvba.apps.excel._refresh import LoadTarget

_ROW = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.DOTALL)
_CELL = re.compile(r"<c\b[^>]*?(?:/>|>.*?</c>)", re.DOTALL)
_SHEET_DATA = re.compile(r"(<sheetData\b[^>]*?)(/>|>(.*?)</sheetData>)", re.DOTALL)
_VALUE = re.compile(r"<v[^>]*>(.*?)</v>", re.DOTALL)
_FORMULA = re.compile(r"<f\b([^>]*)(?:/>|>(.*?)</f>)", re.DOTALL)
_INLINE = re.compile(r"<is>(.*?)</is>", re.DOTALL)
_TEXT = re.compile(r"<t[^>]*>(.*?)</t>", re.DOTALL)
_SHARED_ITEM = re.compile(r"<si>(.*?)</si>", re.DOTALL)
_DEFINED_NAME = re.compile(r"<definedName\b([^>]*)>(.*?)</definedName>", re.DOTALL)
_NUM_FMT = re.compile(r"<numFmt\b([^>]*)/>")
_CELL_XFS = re.compile(r"(<cellXfs\b[^>]*count=\")(\d+)(\"[^>]*>)(.*?)(</cellXfs>)", re.DOTALL)
_XF = re.compile(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", re.DOTALL)

#: The number formats Excel builds in that mean a date or a time.
_BUILTIN_DATE_FORMATS = frozenset({14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47})

_BUILTIN_FORMAT_CODES: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    14: "m/d/yyyy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yyyy h:mm",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    49: "@",
}


class WorkbookFileError(PyOpenVBAError):
    """Raised when a workbook's parts cannot be read or written."""


# --- reading ------------------------------------------------------------------------


def load_workbook(application: Application, path: Path) -> Workbook:
    """A workbook read from an xlsx, xlsm, xlsb or xlam package."""
    from pyopenvba.apps.excel._model import Workbook

    if not path.exists():
        raise WorkbookFileError(f"there is no file at {path}")
    if path.suffix.lower() == ".xlsb":
        raise WorkbookFileError(
            "an xlsb workbook keeps its cells in a binary part that pyOpenVBA does not read; "
            "its VBA project is readable through ExcelFile"
        )
    package = OpcFile.parse(path.read_bytes())
    book = Workbook(application, path.name, str(path.parent))
    book.package = package
    workbook_xml = package.read("xl/workbook.xml").decode("utf-8", errors="replace")
    relationships = _relationship_map(package)
    strings = _shared_strings(package)
    styles = _style_formats(package)
    for name, relationship_id in sheet_entries(workbook_xml):
        part = relationships.get(relationship_id, "")
        sheet = book.add_sheet(name)
        sheet.part_name = part
        if part and package.has(part):
            _read_sheet(sheet, package.read(part).decode("utf-8", errors="replace"), strings, styles)
    _read_names(book, workbook_xml)
    _read_queries(book, path)
    book.saved = True
    return book


def _relationship_map(package: OpcFile) -> dict[str, str]:
    """Each workbook relationship id against the part it names."""
    if not package.has("xl/_rels/workbook.xml.rels"):
        return {}
    text = package.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="replace")
    out: dict[str, str] = {}
    for element in re.findall(r"<Relationship\b[^>]*/>", text):
        attributes = _attributes(element)
        target = attributes.get("Target", "")
        if not target:
            continue
        out[attributes.get("Id", "")] = target[1:] if target.startswith("/") else f"xl/{target.lstrip('/')}"
    return out


def _shared_strings(package: OpcFile) -> list[str]:
    if not package.has("xl/sharedStrings.xml"):
        return []
    text = package.read("xl/sharedStrings.xml").decode("utf-8", errors="replace")
    return ["".join(_unescape(piece) for piece in _TEXT.findall(item)) for item in _SHARED_ITEM.findall(text)]


def _style_formats(package: OpcFile) -> list[str]:
    """The number format of every cell style, by style index."""
    if not package.has("xl/styles.xml"):
        return []
    text = package.read("xl/styles.xml").decode("utf-8", errors="replace")
    custom: dict[int, str] = {}
    for element in _NUM_FMT.findall(text):
        attributes = _attributes(f"<numFmt {element}/>")
        try:
            custom[int(attributes.get("numFmtId", "0"))] = _unescape(attributes.get("formatCode", ""))
        except ValueError:
            continue
    match = _CELL_XFS.search(text)
    if not match:
        return []
    out: list[str] = []
    for element in _XF.findall(match.group(4)):
        attributes = _attributes(element)
        try:
            identifier = int(attributes.get("numFmtId", "0"))
        except ValueError:
            identifier = 0
        out.append(custom.get(identifier, _BUILTIN_FORMAT_CODES.get(identifier, "General")))
    return out


def _read_sheet(sheet: Worksheet, xml: str, strings: list[str], styles: list[str]) -> None:
    from pyopenvba.apps.excel._model import Cell

    match = _SHEET_DATA.search(xml)
    if not match or match.group(2) == "/>":
        return
    for row_xml in _ROW.findall(match.group(3) or ""):
        for cell_xml in _CELL.findall(row_xml):
            attributes = _tag_attributes(cell_xml)
            reference = attributes.get("r", "")
            if not reference:
                continue
            column, row = _split_reference(reference)
            kind = attributes.get("t", "")
            style = attributes.get("s", "")
            number_format = "General"
            if style.isdigit() and int(style) < len(styles):
                number_format = styles[int(style)]
            formula = ""
            formula_match = _FORMULA.search(cell_xml)
            if formula_match is not None:
                body = (formula_match.group(2) or "").strip()
                formula = f"={_unescape(body)}" if body else ""
            value = _cell_value(cell_xml, kind, strings, number_format)
            if value is EMPTY and not formula and number_format == "General":
                continue
            cell = Cell(value=value, formula=formula, number_format=number_format)
            sheet.cells_[(row, column)] = cell


def _cell_value(cell_xml: str, kind: str, strings: list[str], number_format: str) -> object:
    if kind == "inlineStr":
        inline = _INLINE.search(cell_xml)
        return "".join(_unescape(piece) for piece in _TEXT.findall(inline.group(1))) if inline else ""
    raw = _VALUE.search(cell_xml)
    if raw is None:
        return EMPTY
    text = _unescape(raw.group(1))
    if kind == "s":
        try:
            return strings[int(text)]
        except (ValueError, IndexError):
            return ""
    if kind == "str":
        return text
    if kind == "b":
        return text.strip() not in ("0", "")
    if kind == "e":
        from pyopenvba.formula._values import ERRORS, ExcelError

        return ERRORS.get(text, ExcelError(text or "#VALUE!"))
    if kind == "d":
        from pyopenvba.interpreter._values import parse_date_text

        parsed = parse_date_text(text.replace("T", " "))
        return parsed if parsed is not None else text
    try:
        number = float(text)
    except ValueError:
        return text
    if is_date_format(number_format):
        return VBADate(number)
    if number.is_integer() and abs(number) <= 2147483647:
        whole = int(number)
        return VBAInt(whole, "Integer" if -32768 <= whole <= 32767 else "Long")
    return number


def is_date_format(code: str) -> bool:
    """Whether a number format makes its cell a Date rather than a number."""
    if code in ("General", "", "@"):
        return False
    body = re.sub(r"\[[^\]]*\]", "", code)
    body = re.sub(r'"[^"]*"', "", body)
    return any(char in body for char in "ymdhs") and "e+" not in body.lower()


def _read_names(book: Workbook, workbook_xml: str) -> None:
    from pyopenvba.apps.excel._model import NameEntry

    for attributes_text, body in _DEFINED_NAME.findall(workbook_xml):
        attributes = _attributes(f"<definedName {attributes_text}>")
        name = attributes.get("name", "")
        if not name or name.startswith("_xlnm"):
            continue
        book.names_.entries.append(
            NameEntry(name, f"={_unescape(body.strip())}", book, attributes=attributes_text.strip())
        )


_TABLE_PART = re.compile(r"<table\b[^>]*/?>")


def read_load_targets(package: OpcFile, book: Workbook) -> dict[str, LoadTarget]:
    """Where each query's rows sit, read from the tables in the package.

    Excel names the table it makes for a query after the query, so the
    table's name is what ties the two together; its ``ref`` is the block
    the rows fill, headers included.
    """
    from pyopenvba.apps.excel._refresh import LoadTarget as Target

    out: dict[str, LoadTarget] = {}
    wanted = {entry.name.lower(): entry.name for entry in book.queries_.entries}
    for sheet in book.sheets_:
        if not sheet.part_name:
            continue
        rels = f"{sheet.part_name.rsplit('/', 1)[0]}/_rels/{sheet.part_name.rsplit('/', 1)[1]}.rels"
        if not package.has(rels):
            continue
        text = package.read(rels).decode("utf-8", errors="replace")
        for element in re.findall(r"<Relationship\b[^>]*/>", text):
            attributes = _attributes(element)
            target = attributes.get("Target", "")
            if "table" not in target.lower():
                continue
            part = target[1:] if target.startswith("/") else f"xl/{target.lstrip('./')}"
            part = part.replace("xl/../", "xl/").replace("/worksheets/../", "/")
            if not package.has(part):
                continue
            table_xml = package.read(part).decode("utf-8", errors="replace")
            found = _TABLE_PART.search(table_xml)
            if found is None:
                continue
            fields = _tag_attributes(found.group(0))
            name = fields.get("displayName") or fields.get("name", "")
            reference = fields.get("ref", "")
            if not name or not reference:
                continue
            key = wanted.get(name.lower()) or wanted.get(name.replace("_", " ").lower())
            if key is None:
                continue
            try:
                area = Area(*_area_of(reference), sheet.name)
            except ValueError:
                continue
            out[key] = Target(sheet.name, area, part, name)
    return out


def _area_of(reference: str) -> tuple[int, int, int, int]:
    from pyopenvba._a1 import parse_area

    area = parse_area(reference)
    return area.top, area.left, area.bottom, area.right


def _read_queries(book: Workbook, path: Path) -> None:
    from pyopenvba.apps.excel._model import QueryEntry

    try:
        from pyopenvba.powerquery import PowerQueryWorkbook
    except ImportError:  # pragma: no cover - the package always ships
        return
    try:
        with PowerQueryWorkbook(path) as queries:
            for name in queries.query_names():
                query = queries.query(name)
                book.queries_.entries.append(
                    QueryEntry(name, query.formula, getattr(query, "description", "") or "")
                )
    except Exception:  # noqa: BLE001 - a workbook with no Power Query is ordinary
        return


def _split_reference(reference: str) -> tuple[int, int]:
    letters = "".join(char for char in reference if char.isalpha())
    digits = "".join(char for char in reference if char.isdigit())
    return column_number(letters), int(digits)


# --- writing --------------------------------------------------------------------------


def save_workbook(book: Workbook, target: Path) -> None:
    """Write the workbook out, patching only the cells the model changed."""
    package = book.package
    if package is None:
        package = _fresh_package(target.suffix.lower())
        book.package = package
        _match_sheets_to_package(book, package)
    for sheet in book.sheets_:
        # A sheet nobody wrote to keeps the bytes it arrived with.
        if not sheet.part_name or not sheet.dirty:
            continue
        original = (
            package.read(sheet.part_name).decode("utf-8", errors="replace")
            if package.has(sheet.part_name)
            else _EMPTY_SHEET
        )
        package.write(sheet.part_name, _patched_sheet(sheet, original, package).encode("utf-8"))
    if book.names_.changed:
        _write_names(book, package)
    _resize_loaded_tables(book, package)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(package.serialize())


_EMPTY_SHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<dimension ref="A1"/><sheetData/></worksheet>'
)


def _fresh_package(suffix: str) -> OpcFile:
    """A package for a workbook that was made here rather than opened.

    Built from the template captured from a freshly Excel-authored file,
    so what comes out opens without a repair prompt.
    """
    from pyopenvba._templates import EMPTY_XLSM_BYTES

    package = OpcFile.parse(EMPTY_XLSM_BYTES)
    if suffix in (".xlsx", ""):
        _make_macro_free(package)
    return package


def _make_macro_free(package: OpcFile) -> None:
    """Turn the macro-enabled template into a plain xlsx."""
    if package.has("xl/vbaProject.bin"):
        package.remove("xl/vbaProject.bin")
    if package.has("xl/_rels/workbook.xml.rels"):
        text = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        text = re.sub(r"<Relationship\b[^>]*vbaProject\.bin[^>]*/>", "", text)
        package.write("xl/_rels/workbook.xml.rels", text.encode("utf-8"))
    if package.has("[Content_Types].xml"):
        text = package.read("[Content_Types].xml").decode("utf-8")
        text = re.sub(r"<Override\b[^>]*vbaProject\.bin[^>]*/>", "", text)
        text = text.replace(
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        )
        package.write("[Content_Types].xml", text.encode("utf-8"))
    if package.has("xl/workbook.xml"):
        text = package.read("xl/workbook.xml").decode("utf-8")
        text = text.replace(' codeName="ThisWorkbook"', "")
        package.write("xl/workbook.xml", text.encode("utf-8"))


def _match_sheets_to_package(book: Workbook, package: OpcFile) -> None:
    """Give every sheet a part, adding parts for the ones the template lacks."""
    workbook_xml = package.read("xl/workbook.xml").decode("utf-8", errors="replace")
    relationships = _relationship_map(package)
    existing = [(name, relationships.get(rid, "")) for name, rid in sheet_entries(workbook_xml)]
    for index, sheet in enumerate(book.sheets_):
        if index < len(existing):
            sheet.part_name = existing[index][1]
        else:
            sheet.part_name = f"xl/worksheets/sheet{index + 1}.xml"
            package.write(sheet.part_name, _EMPTY_SHEET.encode("utf-8"))
            add_content_type(
                package,
                sheet.part_name,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            )
    _rename_sheets_in_workbook(book, package, existing)


def _resize_loaded_tables(book: Workbook, package: OpcFile) -> None:
    """Follow a refreshed query's rows with its table and its name.

    A refresh that brings back more rows than last time leaves the
    table saying it ends where it used to, and Excel then shows a table
    that stops short of its own data.
    """
    targets = getattr(book, "_load_targets", None)
    if not targets:
        return
    for name, target in targets.items():
        if not target.table_part or not package.has(target.table_part):
            continue
        reference = target.area.address(absolute=False)
        text = package.read(target.table_part).decode("utf-8", errors="replace")
        patched = re.sub(r'(<table\b[^>]*?\bref=")[^"]*(")', rf"\1{reference}\2", text, count=1)
        patched = re.sub(
            r'(<autoFilter\b[^>]*?\bref=")[^"]*(")', rf"\1{reference}\2", patched, count=1
        )
        patched = _follow_columns(package, target, patched)
        if patched != text:
            package.write(target.table_part, patched.encode("utf-8"))
        _follow_defined_name(book, package, name, target)


def _follow_columns(package: OpcFile, target: LoadTarget, text: str) -> str:
    """The table's own column list, brought in line with the headers.

    A refresh that brings back different columns leaves the table
    listing the old ones, and Excel then shows headers that do not
    match the cells under them.
    """
    headers = target.headers
    if not headers:
        return text
    columns = "".join(
        f'<tableColumn id="{index + 1}" uniqueName="{index + 1}" name="{_escape(name)}"'
        f' queryTableFieldId="{index + 1}"/>'
        for index, name in enumerate(headers)
    )
    patched = re.sub(
        r"<tableColumns\b[^>]*>.*?</tableColumns>",
        f'<tableColumns count="{len(headers)}">{columns}</tableColumns>',
        text,
        count=1,
        flags=re.DOTALL,
    )
    _follow_query_table(package, target, headers)
    return patched


def _follow_query_table(package: OpcFile, target: LoadTarget, headers: list[str]) -> None:
    """The query table beside it lists the same fields."""
    part = target.table_part
    rels = f"{part.rsplit('/', 1)[0]}/_rels/{part.rsplit('/', 1)[1]}.rels"
    if not part or not package.has(rels):
        return
    text = package.read(rels).decode("utf-8", errors="replace")
    for element in re.findall(r"<Relationship\b[^>]*/>", text):
        found = _attributes(element).get("Target", "")
        if "querytable" not in found.lower():
            continue
        query_part = found[1:] if found.startswith("/") else f"xl/{found.lstrip('./')}"
        query_part = query_part.replace("xl/../", "xl/").replace("/tables/../", "/")
        if not package.has(query_part):
            continue
        body = package.read(query_part).decode("utf-8", errors="replace")
        fields = "".join(
            f'<queryTableField id="{index + 1}" name="{_escape(name)}" tableColumnId="{index + 1}"/>'
            for index, name in enumerate(headers)
        )
        package.write(
            query_part,
            re.sub(
                r"<queryTableFields\b[^>]*>.*?</queryTableFields>",
                f'<queryTableFields count="{len(headers)}">{fields}</queryTableFields>',
                body,
                count=1,
                flags=re.DOTALL,
            ).encode("utf-8"),
        )


def _follow_defined_name(book: Workbook, package: OpcFile, name: str, target: LoadTarget) -> None:
    """The hidden ExternalData name moves with the table it stands for."""
    from pyopenvba._a1 import quote_sheet

    area = target.area
    wanted = f"{quote_sheet(area.sheet)}!{area.address()}"
    changed = False
    for entry in book.names_.entries:
        if not entry.name.startswith("ExternalData_"):
            continue
        if entry.refers_to.lstrip("=").split("!")[0].strip("'") != area.sheet:
            continue
        if entry.refers_to.lstrip("=") != wanted:
            entry.refers_to = f"={wanted}"
            changed = True
    if changed:
        book.names_.changed = True
        _write_names(book, package)


def _rename_sheets_in_workbook(book: Workbook, package: OpcFile, existing: list[tuple[str, str]]) -> None:
    """Put the model's sheet names into the package's workbook part.

    Only the names of the sheets the template already has: adding a
    sheet to the workbook part means a relationship and a content type
    as well, which is why a new workbook starts from a template rather
    than from nothing.
    """
    text = package.read("xl/workbook.xml").decode("utf-8", errors="replace")
    for index, (old, _) in enumerate(existing):
        if index >= len(book.sheets_):
            break
        wanted = book.sheets_[index].name
        if wanted != old:
            text = text.replace(f'name="{_escape(old)}"', f'name="{_escape(wanted)}"', 1)
    package.write("xl/workbook.xml", text.encode("utf-8"))


def _patched_sheet(sheet: Worksheet, original: str, package: OpcFile) -> str:
    """The sheet's XML with the model's cells written into it."""
    match = _SHEET_DATA.search(original)
    if not match:
        raise WorkbookFileError(f"{sheet.part_name} has no sheetData")
    body = "" if match.group(2) == "/>" else (match.group(3) or "")
    rows: dict[int, str] = {}
    order: list[int] = []
    for row_xml in _ROW.findall(body):
        number = _row_number(row_xml)
        rows[number] = row_xml
        order.append(number)
    for (row, column), cell in sorted(sheet.cells_.items()):
        if cell.is_blank():
            continue
        rows[row] = _with_cell(rows.get(row, f'<row r="{row}"></row>'), row, column, cell, sheet, package)
        if row not in order:
            order.append(row)
    for row in list(rows):
        if not any(not cell.is_blank() for (r, _), cell in sheet.cells_.items() if r == row) and not _CELL.search(
            rows[row]
        ):
            continue
        rows[row] = _without_removed_cells(rows[row], row, sheet)
    rebuilt = "".join(rows[number] for number in sorted(order) if number in rows)
    patched = original[: match.start()] + f"<sheetData>{rebuilt}</sheetData>" + original[match.end() :]
    return _with_dimension(patched, sheet)


def _row_number(row_xml: str) -> int:
    try:
        return int(_tag_attributes(row_xml).get("r", "0"))
    except ValueError:
        return 0


def _with_cell(
    row_xml: str, row: int, column: int, cell: Cell, sheet: Worksheet, package: OpcFile
) -> str:
    """One cell written into its row, replacing whatever was there."""
    reference = f"{column_letter(column)}{row}"
    existing = None
    for candidate in _CELL.findall(row_xml):
        if _tag_attributes(candidate).get("r", "") == reference:
            existing = candidate
            break
    style = _tag_attributes(existing).get("s", "") if existing else ""
    style = _style_for(cell, style, package)
    written = _cell_xml(reference, cell, style)
    if existing is not None:
        return row_xml.replace(existing, written, 1)
    return _insert_cell(row_xml, written, column)


def _insert_cell(row_xml: str, written: str, column: int) -> str:
    """Put a cell into a row, keeping the row's cells in column order."""
    opening = row_xml[: row_xml.index(">") + 1]
    if row_xml.endswith("/>"):
        opening = row_xml[:-2] + ">"
        return f"{opening}{written}</row>"
    inner = row_xml[len(opening) : row_xml.rindex("</row>")]
    pieces = _CELL.findall(inner)
    for index, piece in enumerate(pieces):
        reference = _tag_attributes(piece).get("r", "")
        letters = "".join(char for char in reference if char.isalpha())
        if letters and column_number(letters) > column:
            rebuilt = "".join(pieces[:index]) + written + "".join(pieces[index:])
            return f"{opening}{rebuilt}</row>"
    return f"{opening}{inner}{written}</row>"


def _without_removed_cells(row_xml: str, row: int, sheet: Worksheet) -> str:
    """Drop the cells the model no longer holds anything for."""
    out = row_xml
    for candidate in _CELL.findall(row_xml):
        reference = _tag_attributes(candidate).get("r", "")
        if not reference:
            continue
        try:
            column, number = _split_reference(reference)
        except ValueError:
            continue
        if number != row:
            continue
        cell = sheet.cells_.get((row, column))
        if cell is None or cell.is_blank():
            out = out.replace(candidate, "", 1)
    return out


def _cell_xml(reference: str, cell: Cell, style: str) -> str:
    attributes = f' s="{style}"' if style else ""
    if cell.formula:
        # A formula cell carries its last value in a <v>, never as an
        # inline string: a formula whose answer is text is t="str".
        kind, body = _formula_value(cell)
        return f'<c r="{reference}"{attributes}{kind}><f>{_escape(cell.formula[1:])}</f>{body}</c>'
    kind, body = _value_body(cell.value)
    if not body:
        return f'<c r="{reference}"{attributes}/>'
    return f'<c r="{reference}"{attributes}{kind}>{body}</c>'


def _formula_value(cell: Cell) -> tuple[str, str]:
    """The type and the ``<v>`` a formula cell carries beside its formula."""
    from pyopenvba.formula._values import ExcelError

    if cell.stale or cell.value is EMPTY:
        return "", ""
    value = cell.value
    if isinstance(value, ExcelError):
        return ' t="e"', f"<v>{_escape(value.name)}</v>"
    if isinstance(value, bool):
        return ' t="b"', f"<v>{1 if value else 0}</v>"
    if isinstance(value, str):
        return ' t="str"', f"<v>{_escape(value)}</v>"
    if isinstance(value, VBADate):
        return "", f"<v>{_number_text(value.serial)}</v>"
    if isinstance(value, (int, float)):
        return "", f"<v>{_number_text(value)}</v>"
    return ' t="str"', f"<v>{_escape(to_text(value))}</v>"


def _value_body(value: object) -> tuple[str, str]:
    from pyopenvba.formula._values import ExcelError

    if value is EMPTY:
        return "", ""
    if isinstance(value, ExcelError):
        return ' t="e"', f"<v>{_escape(value.name)}</v>"
    if isinstance(value, bool):
        return ' t="b"', f"<v>{1 if value else 0}</v>"
    if isinstance(value, VBADate):
        return "", f"<v>{_number_text(value.serial)}</v>"
    if isinstance(value, (int, float)):
        return "", f"<v>{_number_text(value)}</v>"
    return ' t="inlineStr"', f"<is><t xml:space=\"preserve\">{_escape(to_text(value))}</t></is>"


def _number_text(value: float) -> str:
    if isinstance(value, int) or float(value).is_integer():
        return str(int(value))
    return repr(float(value))


def _style_for(cell: Cell, existing: str, package: OpcFile) -> str:
    """The style index a cell needs for its number format."""
    wanted = cell.number_format
    if wanted in ("General", ""):
        return existing
    return _ensure_style(package, wanted)


def _ensure_style(package: OpcFile, code: str) -> str:
    """A cellXfs index whose number format is ``code``, adding one if needed."""
    if not package.has("xl/styles.xml"):
        return ""
    text = package.read("xl/styles.xml").decode("utf-8", errors="replace")
    formats = _style_formats(package)
    for index, found in enumerate(formats):
        if found == code:
            return str(index)
    identifier = 0
    for builtin, builtin_code in _BUILTIN_FORMAT_CODES.items():
        if builtin_code == code:
            identifier = builtin
            break
    if not identifier:
        used = [
            int(_attributes(f"<numFmt {one}/>").get("numFmtId", "0"))
            for one in _NUM_FMT.findall(text)
        ]
        identifier = max([163, *used]) + 1
        entry = f'<numFmt numFmtId="{identifier}" formatCode="{_escape(code)}"/>'
        if "<numFmts" in text:
            text = re.sub(
                r'(<numFmts\b[^>]*count=")(\d+)(")',
                lambda match: f"{match.group(1)}{int(match.group(2)) + 1}{match.group(3)}",
                text,
                count=1,
            )
            text = text.replace("</numFmts>", f"{entry}</numFmts>", 1)
        else:
            text = re.sub(
                r"(<styleSheet\b[^>]*>)",
                rf'\1<numFmts count="1">{entry}</numFmts>',
                text,
                count=1,
            )
    match = _CELL_XFS.search(text)
    if not match:
        return ""
    count = int(match.group(2))
    rebuilt = (
        match.group(1)
        + str(count + 1)
        + match.group(3)
        + match.group(4)
        + f'<xf numFmtId="{identifier}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        + match.group(5)
    )
    text = text[: match.start()] + rebuilt + text[match.end() :]
    package.write("xl/styles.xml", text.encode("utf-8"))
    return str(count)


def _with_dimension(xml: str, sheet: Worksheet) -> str:
    bounds = sheet.used_bounds()
    reference = "A1" if bounds is None else Area(*bounds, sheet.name).address(absolute=False)
    return re.sub(r'<dimension\b[^>]*/>', f'<dimension ref="{reference}"/>', xml, count=1)


def _write_names(book: Workbook, package: OpcFile) -> None:
    """The model's defined names, put back into the workbook part."""
    if not package.has("xl/workbook.xml"):
        return
    text = package.read("xl/workbook.xml").decode("utf-8", errors="replace")
    keep = [
        f"<definedName {attributes}>{body}</definedName>"
        for attributes, body in _DEFINED_NAME.findall(text)
        if _attributes(f"<definedName {attributes}>").get("name", "").startswith("_xlnm")
    ]
    for entry in book.names_.entries:
        # A name that came from the file keeps the attributes it came
        # with; localSheetId and hidden are not the model's to drop.
        attributes = entry.attributes or f'name="{_escape(entry.name)}"'
        keep.append(f"<definedName {attributes}>{_escape(entry.refers_to.lstrip('='))}</definedName>")
    block = f"<definedNames>{''.join(keep)}</definedNames>" if keep else ""
    if "<definedNames>" in text:
        text = re.sub(r"<definedNames>.*?</definedNames>", block, text, count=1, flags=re.DOTALL)
    elif "<definedNames/>" in text:
        text = text.replace("<definedNames/>", block, 1)
    elif block:
        text = re.sub(r"(<sheets>.*?</sheets>)", rf"\1{block}", text, count=1, flags=re.DOTALL)
    package.write("xl/workbook.xml", text.encode("utf-8"))
