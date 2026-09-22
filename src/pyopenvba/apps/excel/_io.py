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

import posixpath
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
    from pyopenvba.shapes._values import Shape as ShapeState

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
            sheet_xml = package.read(part).decode("utf-8", errors="replace")
            _read_sheet(sheet, sheet_xml, strings, styles)
            _read_shapes(sheet, package, sheet_xml)
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


def _read_shapes(sheet: Worksheet, package: OpcFile, sheet_xml: str) -> None:
    """The sheet's drawing, read into the model.

    A sheet points at its drawing through a relationship; the shapes
    are in that part, and a form control's own settings are in the
    ``ctrlProps`` part the sheet points at separately.
    """
    from pyopenvba.shapes._xlsx import control_text, grid_of, read_controls, read_drawing

    part = _sheet_relationship(package, sheet.part_name, "drawing")
    if not part or not package.has(part):
        return
    sheet.drawing_part = part
    sheet.drawing_xml = package.read(part).decode("utf-8", errors="replace")
    sheet.shapes_ = read_drawing(sheet.drawing_xml, grid_of(sheet_xml))
    # The counter a new shape's name comes from carries on from what is
    # already there rather than starting again at one.
    sheet.shape_count = len(sheet.shapes_)
    parts = _control_parts(package, sheet.part_name)
    controls = read_controls(sheet_xml, parts)
    vml_part = _sheet_relationship(package, sheet.part_name, "vmlDrawing")
    vml = package.read(vml_part).decode("utf-8") if vml_part and package.has(vml_part) else ""
    for shape in sheet.shapes_:
        found = controls.get(shape.shape_id)
        if found is not None:
            shape.control = found
            # Which part it came from, so a save changes that control
            # rather than adding another one beside it.
            found.part_name = _part_for(package, sheet.part_name, found.relationship)
            # Excel reports a control's macro workbook-qualified, and
            # writes it as [0]!Name; the model carries the plain name.
            shape.macro = found.macro_text.rpartition("!")[2]
            if vml and "<xdr:txBody" not in shape.source:
                shape.text = control_text(vml, shape.shape_id)
    from pyopenvba.apps.excel._radios import initialize

    initialize(sheet)
    sheet.drawing_dirty = False


def _sheet_relationship(package: OpcFile, sheet_part: str, kind: str) -> str:
    """The part a sheet's relationship of this kind points at.

    Matched on the relationship's type rather than its target, because
    xl/drawings holds the VML for the form controls as well, and its
    name starts with the same word.
    """
    for _, (relationship_kind, target) in _sheet_relationships(package, sheet_part).items():
        if relationship_kind == kind:
            return _resolved(target)
    return ""


def _sheet_relationships(package: OpcFile, sheet_part: str) -> dict[str, tuple[str, str]]:
    """Each relationship on a sheet, as ``id -> (kind, target)``.

    The kind is the last word of the type URI: drawing, vmlDrawing,
    ctrlProp, table, and so on.
    """
    if not sheet_part:
        return {}
    folder, _, name = sheet_part.rpartition("/")
    rels = f"{folder}/_rels/{name}.rels"
    if not package.has(rels):
        return {}
    text = package.read(rels).decode("utf-8", errors="replace")
    out: dict[str, tuple[str, str]] = {}
    for element in re.findall(r"<Relationship\b[^>]*/>", text):
        fields = _attributes(element)
        if "Id" in fields and "Target" in fields:
            out[fields["Id"]] = (fields.get("Type", "").rpartition("/")[2], fields["Target"])
    return out


def _resolved(target: str) -> str:
    """A relationship target as a part name in the package."""
    if target.startswith("/"):
        return target[1:]
    part = f"xl/worksheets/{target}"
    while "/../" in part:
        head, _, tail = part.rpartition("/../")
        part = head.rpartition("/")[0] + "/" + tail
    return part.replace("./", "")


def _part_for(package: OpcFile, sheet_part: str, relationship: str) -> str:
    """The part one of a sheet's relationships names."""
    found = _sheet_relationships(package, sheet_part).get(relationship)
    return _resolved(found[1]) if found else ""


def _control_parts(package: OpcFile, sheet_part: str) -> dict[str, str]:
    """Each control's own part, by the relationship the sheet names."""
    out: dict[str, str] = {}
    for relationship, (kind, target) in _sheet_relationships(package, sheet_part).items():
        if kind != "ctrlProp":
            continue
        part = _resolved(target)
        if package.has(part):
            out[relationship] = package.read(part).decode("utf-8", errors="replace")
    return out


def _write_shapes(book: Workbook, package: OpcFile) -> None:
    """Write back the drawing of every sheet whose shapes changed."""
    from pyopenvba.shapes._xlsx import EMPTY_DRAWING, read_drawing, written

    for sheet in book.sheets_:
        if not sheet.drawing_dirty:
            continue
        part = sheet.drawing_part or _new_drawing(book, package, sheet)
        if not part:
            continue
        original = sheet.drawing_xml or EMPTY_DRAWING
        before = read_drawing(original, sheet.drawing_grid())
        package.write(part, written(sheet.shapes_, original, sheet.drawing_grid()).encode("utf-8"))
        sheet.drawing_part = part
        sheet.drawing_xml = package.read(part).decode("utf-8", errors="replace")
        _sync_controls(package, sheet, before)
        _write_new_controls(package, sheet)
        _write_control_macros(package, sheet)
        sheet.drawing_dirty = False


def _sync_controls(package: OpcFile, sheet: Worksheet, before: list[ShapeState]) -> None:
    """Follow control edits and deletions into the sheet, VML and part relationships."""
    from pyopenvba.shapes._xlsx import (
        control_text, read_controls, with_control_shape, with_vml_control,
        without_control, without_vml_control, with_control_bindings, with_vml_bindings,
    )

    if not sheet.part_name or not package.has(sheet.part_name):
        return
    xml = package.read(sheet.part_name).decode("utf-8")
    controls = read_controls(xml, _control_parts(package, sheet.part_name))
    vml_part = _sheet_relationship(package, sheet.part_name, "vmlDrawing")
    vml = package.read(vml_part).decode("utf-8") if vml_part and package.has(vml_part) else ""
    current = {item.shape_id: item for item in sheet.shapes_}
    for old in before:
        info = controls.get(old.shape_id)
        if old.kind != "formControl" or info is None:
            continue
        item = current.get(old.shape_id)
        if item is None or item.control is None or item.control.relationship != info.relationship:
            part = _part_for(package, sheet.part_name, info.relationship)
            xml = without_control(xml, old.shape_id)
            vml = without_vml_control(vml, old.shape_id)
            _remove_sheet_relationship(package, sheet.part_name, info.relationship)
            if part:
                _remove_unreferenced_part(package, part)
            continue
        if "<xdr:txBody" not in old.source:
            old.text = control_text(vml, old.shape_id)
        if (item.name, item.left, item.top, item.width, item.height) != (
            old.name, old.left, old.top, old.width, old.height,
        ):
            xml = with_control_shape(xml, item, sheet.drawing_grid())
        vml = with_vml_control(vml, item, old, sheet.drawing_grid())
        if any(getattr(item.control, key) != getattr(info, key) for key in (
            "linked_cell", "list_range", "value", "selection_mode", "items", "selected_indices",
            "minimum", "maximum", "increment", "page_change", "drop_width", "first_button",
        )):
            part = _part_for(package, sheet.part_name, info.relationship)
            if part and package.has(part):
                properties = package.read(part).decode("utf-8")
                package.write(part, with_control_bindings(properties, item.control).encode("utf-8"))
            vml = with_vml_bindings(vml, item.shape_id, item.control)
    if vml_part:
        if re.search(r"<v:shape(?=[\s/>])", vml):
            package.write(vml_part, vml.encode("utf-8"))
        else:
            for rid, (kind, _) in _sheet_relationships(package, sheet.part_name).items():
                if kind == "vmlDrawing":
                    _remove_sheet_relationship(package, sheet.part_name, rid)
                    xml = re.sub(
                        r"<legacyDrawing\b[^>]*/>",
                        lambda m: "" if _attributes(m.group(0)).get("r:id") == rid else m.group(0), xml,
                    )
            _remove_unreferenced_part(package, vml_part)
    package.write(sheet.part_name, xml.encode("utf-8"))


def _remove_sheet_relationship(package: OpcFile, sheet_part: str, relationship: str) -> None:
    folder, _, name = sheet_part.rpartition("/")
    part = f"{folder}/_rels/{name}.rels"
    if package.has(part):
        text = package.read(part).decode("utf-8")
        text = re.sub(r"<Relationship\b[^>]*/>",
                      lambda m: "" if _attributes(m.group(0)).get("Id") == relationship else m.group(0), text)
        package.write(part, text.encode("utf-8"))


def _remove_unreferenced_part(package: OpcFile, part: str) -> None:
    """A control part may be shared; remove it only after its final relationship."""
    for name in package.names():
        if not name.endswith(".rels"):
            continue
        folder = name.split("/_rels/", 1)[0] if "/_rels/" in name else ""
        for element in re.findall(r"<Relationship\b[^>]*/>", package.read(name).decode("utf-8")):
            attrs = _attributes(element)
            if attrs.get("TargetMode") == "External":
                continue
            target = attrs.get("Target", "")
            resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
            if resolved == part:
                return
    package.remove(part)
    types = package.read("[Content_Types].xml").decode("utf-8")
    types = re.sub(r"<Override\b[^>]*/>",
                   lambda m: "" if _attributes(m.group(0)).get("PartName") == f"/{part}" else m.group(0), types)
    package.write("[Content_Types].xml", types.encode("utf-8"))


def _write_new_controls(package: OpcFile, sheet: Worksheet) -> None:
    """Make real controls of the form controls a macro added.

    A form control is four things at once: the hidden shape in the
    drawing, the sheet's own ``<control>`` record, a part of its own
    saying what kind it is, and the VML Excel actually draws.  Miss any
    of them and Excel either loses the control or asks to repair the
    file.
    """
    from pyopenvba.shapes._xlsx import EMPTY_VML, control_entry, control_vml, with_vml_shape

    fresh = [
        one
        for one in sheet.shapes_
        if one.kind == "formControl" and one.control is not None and not one.control.part_name
    ]
    if not fresh or not sheet.part_name or not package.has(sheet.part_name):
        return
    sheet_xml = package.read(sheet.part_name).decode("utf-8", errors="replace")
    vml_part = _sheet_relationship(package, sheet.part_name, "vmlDrawing")
    if not vml_part:
        vml_part = _new_vml(package, sheet)
        sheet_xml = package.read(sheet.part_name).decode("utf-8", errors="replace")
    vml = package.read(vml_part).decode("utf-8", errors="replace") if package.has(vml_part) else EMPTY_VML
    grid = sheet.drawing_grid()
    entries: list[str] = []
    for shape in fresh:
        part = _new_control_part(package, shape)
        relationship = _add_sheet_relationship(
            package,
            sheet.part_name,
            f"../ctrlProps/{part.rsplit('/', 1)[1]}",
            # The type Excel writes, read out of a file it saved.  A
            # control whose r:id resolves to nothing makes a workbook
            # Excel will not open at all.
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp",
        )
        if shape.control is not None:
            shape.control.part_name = part
            shape.control.relationship = relationship
        entries.append(control_entry(shape, relationship, grid))
        vml = with_vml_shape(vml, control_vml(shape, grid))
    package.write(vml_part, vml.encode("utf-8"))
    package.write(sheet.part_name, _with_controls(sheet_xml, entries).encode("utf-8"))


def _new_control_part(package: OpcFile, shape: ShapeState) -> str:
    """A control's own part, written where Excel keeps them."""
    from pyopenvba.shapes._xlsx import control_properties

    taken = {name for name in package.names() if name.startswith("xl/ctrlProps/ctrlProp")}
    number = 1
    while f"xl/ctrlProps/ctrlProp{number}.xml" in taken:
        number += 1
    part = f"xl/ctrlProps/ctrlProp{number}.xml"
    package.write(part, control_properties(shape).encode("utf-8"))
    add_content_type(package, part, "application/vnd.ms-excel.controlproperties+xml")
    return part


def _new_vml(package: OpcFile, sheet: Worksheet) -> str:
    """The VML part a sheet needs before it can hold a control."""
    from pyopenvba.shapes._xlsx import EMPTY_VML

    taken = {name for name in package.names() if name.startswith("xl/drawings/vmlDrawing")}
    number = 1
    while f"xl/drawings/vmlDrawing{number}.vml" in taken:
        number += 1
    part = f"xl/drawings/vmlDrawing{number}.vml"
    package.write(part, EMPTY_VML.encode("utf-8"))
    _add_default_content_type(package, "vml", "application/vnd.openxmlformats-officedocument.vmlDrawing")
    relationship = _add_sheet_relationship(
        package,
        sheet.part_name,
        f"../drawings/vmlDrawing{number}.vml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
    )
    text = package.read(sheet.part_name).decode("utf-8", errors="replace")
    if "<legacyDrawing " not in text:
        element = f'<legacyDrawing r:id="{relationship}"/>'
        # After the drawing, which is where Excel puts it.
        if "<drawing " in text:
            at = text.find("/>", text.find("<drawing ")) + 2
            text = text[:at] + element + text[at:]
        else:
            text = text.replace("</worksheet>", element + "</worksheet>")
        package.write(sheet.part_name, text.encode("utf-8"))
    return part


def _add_default_content_type(package: OpcFile, extension: str, content_type: str) -> None:
    """One more Default in [Content_Types], if it is not there already.

    It goes with the other Defaults, before the first Override: the
    package schema wants them in that order, and Excel refuses to open
    a workbook whose content types are out of it.
    """
    raw = package.read("[Content_Types].xml").decode("utf-8")
    if f'Extension="{extension}"' in raw:
        return
    element = f'<Default Extension="{extension}" ContentType="{content_type}"/>'
    at = raw.rfind("/>", 0, raw.find("<Override"))
    if at < 0:
        at = raw.find(">", raw.find("<Types"))
    package.write("[Content_Types].xml", (raw[: at + 2] + element + raw[at + 2 :]).encode("utf-8"))


#: What a control's markup needs declared on the worksheet element: the
#: drawing prefix its anchor uses, and the 2009 prefix the
#: AlternateContent requires.  A sheet Excel wrote with controls has
#: both; one written from the template has neither, and leaving them
#: out makes a file Excel will not open.
_CONTROL_NAMESPACES = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
}


def _with_namespaces(sheet_xml: str, wanted: dict[str, str]) -> str:
    """The sheet with these prefixes declared, if they are not already."""
    found = re.search(r"<worksheet\b[^>]*>", sheet_xml)
    if found is None:
        return sheet_xml
    element = found.group(0)
    additions = "".join(
        f' xmlns:{prefix}="{uri}"'
        for prefix, uri in wanted.items()
        if f"xmlns:{prefix}=" not in element
    )
    if not additions:
        return sheet_xml
    changed = element[:-1].rstrip() + additions + ">"
    return sheet_xml[: found.start()] + changed + sheet_xml[found.end() :]


def _with_controls(sheet_xml: str, entries: list[str]) -> str:
    """The sheet carrying these control records, where Excel keeps them."""
    if not entries:
        return sheet_xml
    sheet_xml = _with_namespaces(sheet_xml, _CONTROL_NAMESPACES)
    body = "".join(entries)
    marker = "<controls>"
    at = sheet_xml.find(marker)
    if at >= 0:
        head = at + len(marker)
        return sheet_xml[:head] + body + sheet_xml[head:]
    block = (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="x14">'
        f"<controls>{body}</controls>"
        "</mc:Choice>"
        "</mc:AlternateContent>"
    )
    return sheet_xml.replace("</worksheet>", block + "</worksheet>")


def _write_control_macros(package: OpcFile, sheet: Worksheet) -> None:
    """Follow a form control's macro into the two parts that hold it.

    A control's macro is not in the drawing at all: the sheet's own
    ``controlPr`` carries it, and so does the ``x:FmlaMacro`` of the
    VML shape Excel reads.  Written as Excel writes it, ``[0]!Name``
    for a procedure in this workbook.
    """
    from pyopenvba.shapes._xlsx import control_macro, with_control_macro, with_vml_macro

    controls = [one for one in sheet.shapes_ if one.kind == "formControl"]
    if not controls or not sheet.part_name or not package.has(sheet.part_name):
        return
    sheet_xml = package.read(sheet.part_name).decode("utf-8", errors="replace")
    vml_part = _sheet_relationship(package, sheet.part_name, "vmlDrawing")
    vml = package.read(vml_part).decode("utf-8", errors="replace") if vml_part else ""
    changed_sheet = False
    changed_vml = False
    for shape in controls:
        wanted = f"[0]!{shape.macro}" if shape.macro else ""
        if control_macro(sheet_xml, shape.shape_id) == wanted:
            continue
        sheet_xml = with_control_macro(sheet_xml, shape.shape_id, wanted)
        changed_sheet = True
        if vml:
            vml = with_vml_macro(vml, shape.shape_id, wanted)
            changed_vml = True
    if changed_sheet:
        package.write(sheet.part_name, sheet_xml.encode("utf-8"))
    if changed_vml:
        package.write(vml_part, vml.encode("utf-8"))


def _new_drawing(book: Workbook, package: OpcFile, sheet: Worksheet) -> str:
    """A drawing part for a sheet that had none, wired to the sheet.

    Three things make a drawing part real to Excel: the part itself,
    the content type that says what it is, and the sheet's own
    ``<drawing>`` element naming the relationship to it.
    """
    from pyopenvba.shapes._xlsx import EMPTY_DRAWING

    if not sheet.part_name:
        return ""
    taken = {
        name for name in package.names() if name.startswith("xl/drawings/drawing")
    }
    number = 1
    while f"xl/drawings/drawing{number}.xml" in taken:
        number += 1
    part = f"xl/drawings/drawing{number}.xml"
    package.write(part, EMPTY_DRAWING.encode("utf-8"))
    add_content_type(
        package,
        part,
        "application/vnd.openxmlformats-officedocument.drawing+xml",
    )
    relationship = _add_sheet_relationship(
        package,
        sheet.part_name,
        f"../drawings/drawing{number}.xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    )
    _name_the_drawing(package, sheet, relationship)
    return part


def _add_sheet_relationship(package: OpcFile, sheet_part: str, target: str, kind: str) -> str:
    """One more relationship on a sheet, and the id it was given."""
    folder, _, name = sheet_part.rpartition("/")
    rels = f"{folder}/_rels/{name}.rels"
    if package.has(rels):
        text = package.read(rels).decode("utf-8", errors="replace")
    else:
        text = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            "</Relationships>"
        )
    used = {
        fields.get("Id", "")
        for fields in map(_attributes, re.findall(r"<Relationship\b[^>]*/>", text))
    }
    number = 1
    while f"rId{number}" in used:
        number += 1
    identifier = f"rId{number}"
    element = f'<Relationship Id="{identifier}" Type="{kind}" Target="{target}"/>'
    package.write(rels, text.replace("</Relationships>", element + "</Relationships>").encode("utf-8"))
    return identifier


def _name_the_drawing(package: OpcFile, sheet: Worksheet, relationship: str) -> None:
    """Put the ``<drawing>`` element in the sheet, where Excel keeps it.

    It goes at the end of the sheet, after everything else Excel writes;
    a sheet that already names a drawing is left alone.
    """
    text = package.read(sheet.part_name).decode("utf-8", errors="replace")
    if "<drawing " in text:
        return
    element = f'<drawing r:id="{relationship}"/>'
    if "</worksheet>" not in text:
        return
    package.write(sheet.part_name, text.replace("</worksheet>", element + "</worksheet>").encode("utf-8"))


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
    from pyopenvba.apps.excel._controls import refresh

    for sheet in book.sheets_:
        for shape in sheet.shapes_:
            refresh(sheet, shape)
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
    _write_shapes(book, package)
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
    from pyopenvba.interpreter._values import VBAErrorValue

    if isinstance(value, VBAErrorValue):
        from pyopenvba.apps.excel._calc import ERROR_NUMBERS

        value = ExcelError(next((name for name, number in ERROR_NUMBERS.items()
                                if number == value.number), "#VALUE!"))

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
        from pyopenvba._a1 import split_sheet

        owned = _attributes(f"<definedName {entry.attributes}>")
        scope, bare = split_sheet(entry.name)
        owned["name"] = bare
        if scope:
            index = next((i for i, sheet in enumerate(book.sheets_) if sheet.name.lower() == scope.lower()), None)
            if index is not None:
                owned["localSheetId"] = str(index)
        attributes = " ".join(f'{key}="{_escape(value)}"' for key, value in owned.items())
        keep.append(f"<definedName {attributes}>{_escape(entry.refers_to.lstrip('='))}</definedName>")
    block = f"<definedNames>{''.join(keep)}</definedNames>" if keep else ""
    if "<definedNames>" in text:
        text = re.sub(r"<definedNames>.*?</definedNames>", block, text, count=1, flags=re.DOTALL)
    elif "<definedNames/>" in text:
        text = text.replace("<definedNames/>", block, 1)
    elif block:
        text = re.sub(r"(<sheets>.*?</sheets>)", rf"\1{block}", text, count=1, flags=re.DOTALL)
    package.write("xl/workbook.xml", text.encode("utf-8"))
