"""The same book as scripts/measure_tool_books.py, spelled another way: each vocabulary under another prefix or as the default,
and elements a writer may leave out for their defaults. Excel and the model each read every member of every variant, and what
Excel makes of a variant it will not open is recorded too.

    python scripts/measure_tool_book_spellings.py

writes tests/fixtures/tool_books/spellings.json, which tests/test_excel_tool_book_spellings.py replays.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_tool_books import OUT, READS, SETUP, model_reads, source, tally  # noqa: E402

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
CORE = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DUBLIN = "http://purl.org/dc/elements/1.1/"
APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
#: How each vocabulary is spelled where a variant does not change it.
USUAL = {MAIN: "", OFFICE_RELATIONSHIPS: "r", PACKAGE_RELATIONSHIPS: "", CONTENT_TYPES: "", DRAWING: "xdr", DRAWINGML: "a", CHART: "c", CORE: "cp", DUBLIN: "dc",
         "http://purl.org/dc/terms/": "dcterms", "http://purl.org/dc/dcmitype/": "dcmitype", "http://www.w3.org/2001/XMLSchema-instance": "xsi", APP: "",
         "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes": "vt"}
Parts = dict[str, bytes]


def spelled(data: bytes, wanted: dict[str, str]) -> bytes:
    """A part with its vocabularies under the prefixes ``wanted`` gives, "" for the default namespace."""
    prefixes = {**USUAL, **wanted}
    root = ET.fromstring(data)
    default = next((uri for uri, prefix in prefixes.items() if prefix == "" and root.tag.startswith("{" + uri + "}")), None)
    for uri, prefix in prefixes.items():
        if prefix:
            ET.register_namespace(prefix, uri)
    for uri in prefixes:
        if prefixes[uri] == "" and uri != default:
            ET.register_namespace("z" + str(sorted(prefixes).index(uri)), uri)
    if default is not None:
        ET.register_namespace("zdefault", default)
    text = ET.tostring(root, encoding="unicode").replace("xmlns:zdefault=", "xmlns=")
    text = re.sub(r"(</?)zdefault:", r"\1", text)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + text).encode("utf-8")


def under(uri: str, prefix: str, parts_like: str = r"\.(xml|rels)$") -> Callable[[Parts], Parts]:
    def change(parts: Parts) -> Parts:
        made = dict(parts)
        for name, data in parts.items():
            if re.search(parts_like, name) and uri.encode() in data:
                made[name] = spelled(data, {uri: prefix})
        return made
    return change


def replaced(part_like: str, pattern: str, new: str = "") -> Callable[[Parts], Parts]:
    """A variant taking a pattern out of the parts it names."""
    def change(parts: Parts) -> Parts:
        made = dict(parts)
        for name, data in parts.items():
            if re.search(part_like, name):
                made[name] = re.sub(pattern.encode(), new.encode(), data, flags=re.S)
        return made
    return change


def without_part(part_like: str, rels_target: str) -> Callable[[Parts], Parts]:
    """A variant with a part gone, and what names it: its relationship and content type."""
    def change(parts: Parts) -> Parts:
        made = {name: data for name, data in parts.items() if not re.search(part_like, name)}
        for name, data in list(made.items()):
            if name.endswith(".rels"):
                made[name] = re.sub(rb"<Relationship [^>]*(?:" + rels_target.encode() + rb")[^>]*/>", b"", data)
            elif name == "[Content_Types].xml":
                made[name] = re.sub(rb"<Override [^>]*(?:" + rels_target.encode() + rb")[^>]*/>", b"", data)
        return made
    return change


def vml_under(v: str, o: str, x: str) -> Callable[[Parts], Parts]:
    def change(parts: Parts) -> Parts:
        made = dict(parts)
        for name, data in parts.items():
            if name.endswith(".vml"):
                text = data.decode()
                for old, new in (("v", v), ("o", o), ("x", x)):
                    text = re.sub(rf"(</?){old}:", rf"\1{new}:", text)
                    text = re.sub(rf"(\s){old}:([\w.-]+=)", rf"\1{new}:\2", text)
                    text = text.replace(f"xmlns:{old}=", f"xmlns:{new}=")
                made[name] = text.encode()
        return made
    return change


SHEET = r"xl/worksheets/sheet\d+\.xml$"
FIRST_SHEET = r"xl/worksheets/sheet1\.xml$"
STYLES = r"xl/styles\.xml$"
BOOK = r"xl/workbook\.xml$"
#: Each variant, what a writer may do that Excel's own save never does.
VARIANTS: dict[str, Callable[[Parts], Parts]] = {
    # Prefixes
    "main_prefixed": under(MAIN, "s", r"^xl/[^/]*(/[^/]*)?\.xml$"),
    "relationship_ids_prefixed_rel": under(OFFICE_RELATIONSHIPS, "rel", r"\.xml$"),
    "rels_prefixed": under(PACKAGE_RELATIONSHIPS, "p", r"\.rels$"),
    "content_types_prefixed": under(CONTENT_TYPES, "t", r"^\[Content_Types\]\.xml$"),
    "drawing_default": under(DRAWING, "", r"xl/drawings/drawing\d+\.xml$"),
    "drawing_prefixed_d": under(DRAWING, "d", r"xl/drawings/drawing\d+\.xml$"),
    "drawingml_prefixed_g": under(DRAWINGML, "g"),
    "chart_default": under(CHART, "", r"xl/charts/chart\d+\.xml$"),
    "chart_prefixed_ch": under(CHART, "ch"),
    "core_prefixed": lambda parts: under(DUBLIN, "e")(under(CORE, "k", r"docProps/core\.xml$")(parts)),
    "app_prefixed": under(APP, "q", r"docProps/app\.xml$"),
    "vml_prefixed": vml_under("vv", "oo", "xx"),
    # Left out
    "no_cell_addresses": replaced(FIRST_SHEET, r'(<c) r="[A-Z]+\d+"', r"\1"),
    "no_cell_addresses_by_a_table": replaced(SHEET, r'(<c) r="[A-Z]+\d+"', r"\1"),
    "no_row_numbers": replaced(FIRST_SHEET, r'(<row) r="\d+"', r"\1"),
    "no_row_numbers_by_a_table": replaced(SHEET, r'(<row) r="\d+"', r"\1"),
    "no_row_or_cell_addresses": lambda parts: replaced(FIRST_SHEET, r'(<c) r="[A-Z]+\d+"', r"\1")(replaced(FIRST_SHEET, r'(<row) r="\d+"', r"\1")(parts)),
    "no_spans": replaced(SHEET, r' spans="[^"]*"'),
    "no_dimension": replaced(SHEET, r"<dimension [^>]*/>"),
    "no_sheet_views": replaced(SHEET, r"<sheetViews>.*?</sheetViews>"),
    "no_sheet_format": replaced(SHEET, r"<sheetFormatPr [^>]*/>"),
    "no_page_margins": replaced(SHEET, r"<pageMargins [^>]*/>"),
    "no_cached_values": replaced(SHEET, r"(</f>)<v>[^<]*</v>", r"\1"),
    "no_book_views": replaced(BOOK, r"<bookViews>.*?</bookViews>"),
    "no_calc_pr": replaced(BOOK, r"<calcPr [^>]*/>"),
    "no_workbook_pr": replaced(BOOK, r"<workbookPr [^>]*/>"),
    "no_file_version": replaced(BOOK, r"<fileVersion [^>]*/>"),
    "no_cell_styles": replaced(STYLES, r"<cellStyles.*?</cellStyles>"),
    "no_cell_style_xfs": replaced(STYLES, r"<cellStyleXfs.*?</cellStyleXfs>"),
    "no_style_counts": replaced(STYLES, r' count="\d+"'),
    "no_dxfs": replaced(STYLES, r"<dxfs [^>]*/>"),
    "no_table_styles": replaced(STYLES, r"<tableStyles [^>]*/>"),
    "no_font_scheme_or_family": replaced(STYLES, r'<(scheme|family) val="[^"]*"/>'),
    "no_theme": without_part(r"xl/theme/", "theme"),
    "no_doc_props": without_part(r"docProps/", "docProps|extended-properties|core-properties"),
    "no_shared_string_counts": replaced(r"xl/sharedStrings\.xml$", r' (count|uniqueCount)="\d+"'),
    "no_chart_caches": replaced(r"xl/charts/chart\d+\.xml$", r"<c:(strCache|numCache)>.*?</c:\1>"),
    "no_table_xr": replaced(r"xl/tables/", r' totalsRowShown="\d+"'),
    "no_comment_shape_ids": replaced(r"xl/comments", r' shapeId="\d+"'),
    "no_note_anchor": replaced(r"\.vml$", r"<x:Anchor>.*?</x:Anchor>"),
}
OPEN = """Function Tried(ByVal file As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(file, UpdateLinks:=0)
    Tried = Err.Number & "|" & Err.Description
    Err.Clear
    wb.Close False
End Function
"""


def made(base: Path, variant: str, file: Path) -> None:
    with zipfile.ZipFile(base) as given:
        parts = {name: given.read(name) for name in given.namelist()}
    changed = VARIANTS[variant](parts)
    with zipfile.ZipFile(file, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in changed.items():
            out.writestr(name, data)


def main() -> None:
    """Named variants, or all of them, put into the record beside what it holds: under load a whole run outlasts a command."""
    from pyvbaharness import ExcelSession, HarnessConfig

    base = OUT / "xlsxwriter.xlsx"
    record_file = OUT / "spellings.json"
    rows: dict[str, object] = json.loads(record_file.read_text(encoding="utf-8"))["variants"] if record_file.exists() else {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_spell_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        for variant in sys.argv[1:] or VARIANTS:
            file = Path(folder) / f"{variant}.xlsx"
            made(base, variant, file)
            opened = excel.run_vba(OPEN, "Tried", (str(file),), timeout=60.0, line_numbers=False)
            answer = str(opened.value) if opened.ok else f"{opened.outcome} {opened.message}"
            theirs: list[str] = []
            if answer.startswith("0|"):
                read = excel.run_vba(source(READS, "typed"), "Go", (str(file),), timeout=120.0, line_numbers=False)
                theirs = str(read.value).split("\t") if read.ok else []
            mine = model_reads(file, READS, "typed")
            rows[variant] = {"opens": answer.replace(file.name, "<book>"), "excel": theirs, "model": mine}
            print(variant, answer[:60], tally(theirs, mine) if theirs else "no reads", "model first:", mine[0][:60], flush=True)
            record = {"setup": SETUP, "variants": rows, "measured_at": datetime.now(timezone.utc).isoformat()}
            record_file.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
