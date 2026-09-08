"""Putting a query's result on a worksheet, and taking it off again.

What the metadata says about where a query loads is not what makes it
load: Excel needs the connection, the query table, the table and the
sheet's reference to it.  These check that all of that is written, and
that unloading takes every piece back out -- including the connections
part itself, which Excel refuses when it holds no connections.
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path
from xml.etree import ElementTree

import pytest

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery import _metadata as meta
from pyopenvba.powerquery import LOAD_CONNECTION_ONLY, LOAD_TABLE, PowerQueryWorkbook
from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.powerquery._sheets import CellRef, column_letter, column_number

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"


@pytest.fixture
def book(tmp_path: Path) -> PowerQueryWorkbook:
    out = tmp_path / "three.xlsx"
    shutil.copyfile(FIXTURES / "three_queries.xlsx", out)
    return PowerQueryWorkbook(out)


@pytest.fixture
def loaded(tmp_path: Path) -> PowerQueryWorkbook:
    out = tmp_path / "loaded.xlsx"
    shutil.copyfile(FIXTURES / "loaded_to_sheet.xlsx", out)
    return PowerQueryWorkbook(out)


def parts_of(book: PowerQueryWorkbook) -> list[str]:
    return book._opc.names()  # pyright: ignore[reportPrivateUsage]


def read(book: PowerQueryWorkbook, part: str) -> str:
    raw = book._opc.read(part)  # pyright: ignore[reportPrivateUsage]
    return raw.decode("utf-16") if raw[:2] == bytes((0xFF, 0xFE)) else raw.decode("utf-8")


# --- references ---------------------------------------------------------------


@pytest.mark.parametrize(("number", "letters"), [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (703, "AAA")])
def test_a_column_number_and_its_letters_agree(number: int, letters: str) -> None:
    assert column_letter(number) == letters
    assert column_number(letters) == number


def test_a_cell_reference_reads_and_writes(  ) -> None:
    assert str(CellRef.parse("H1")) == "H1"
    assert CellRef.parse("$AB$12") == CellRef(28, 12)
    with pytest.raises(PowerQueryError, match="not a cell reference"):
        CellRef.parse("nowhere")


# --- loading ------------------------------------------------------------------


def test_loading_writes_every_object_excel_needs(book: PowerQueryWorkbook) -> None:
    book.add_query("Loadable", "let Source = 1 in Source")
    assert book.load_to_sheet("Loadable", ["A", "B"], cell="H1") == "Loadable"

    names = parts_of(book)
    assert "xl/connections.xml" in names
    assert "xl/tables/table1.xml" in names
    assert "xl/queryTables/queryTable1.xml" in names
    assert "xl/tables/_rels/table1.xml.rels" in names
    assert "xl/worksheets/_rels/sheet1.xml.rels" in names

    connections = read(book, "xl/connections.xml")
    assert "Provider=Microsoft.Mashup.OleDb.1" in connections
    assert "Location=Loadable;" in connections
    assert "SELECT * FROM [Loadable]" in connections

    table = read(book, "xl/tables/table1.xml")
    assert 'ref="H1:I2"' in table
    assert 'tableType="queryTable"' in table
    assert 'name="A"' in table and 'name="B"' in table

    sheet = read(book, "xl/worksheets/sheet1.xml")
    assert "<tableParts count=\"1\">" in sheet
    assert "<is><t>A</t></is>" in sheet and "<is><t>B</t></is>" in sheet
    assert 'r="H1"' in sheet

    types = read(book, "[Content_Types].xml")
    assert "/xl/tables/table1.xml" in types
    assert "/xl/queryTables/queryTable1.xml" in types
    assert "/xl/connections.xml" in types
    assert "ExternalData_1" in read(book, "xl/workbook.xml")


def test_loading_says_so_in_the_metadata(book: PowerQueryWorkbook) -> None:
    book.add_query("Loadable", "let Source = 1 in Source")
    book.load_to_sheet("Loadable", ["A"])
    query = book.query("Loadable")
    assert query.load_target == LOAD_TABLE
    assert query.load_enabled is True
    assert query.target_name == "Loadable"
    entries = query.entries()
    assert entries[meta.FILL_TARGET_NAME_CUSTOMIZED] == 1
    assert entries[meta.NAME_UPDATED_AFTER_FILL] == 0


def test_every_part_a_load_writes_is_well_formed(book: PowerQueryWorkbook) -> None:
    book.add_query("Loadable", "let Source = 1 in Source")
    book.load_to_sheet("Loadable", ["A", "B"], cell="H1")
    for name in parts_of(book):
        if name.endswith((".xml", ".rels")):
            ElementTree.fromstring(read(book, name))


def test_a_table_can_be_given_its_own_name(book: PowerQueryWorkbook) -> None:
    book.add_query("Loadable", "let Source = 1 in Source")
    assert book.load_to_sheet("Loadable", ["A"], table_name="OrderLines") == "OrderLines"
    assert 'displayName="OrderLines"' in read(book, "xl/tables/table1.xml")
    assert book.query("Loadable").target_name == "OrderLines"


def test_a_second_load_gets_its_own_parts_and_connection(book: PowerQueryWorkbook) -> None:
    book.add_query("One", "let Source = 1 in Source")
    book.add_query("Two", "let Source = 2 in Source")
    book.load_to_sheet("One", ["A"], cell="H1")
    book.load_to_sheet("Two", ["B"], cell="K1")
    assert "xl/tables/table2.xml" in parts_of(book)
    assert read(book, "xl/connections.xml").count("<connection ") == 2
    assert '<tableParts count="2">' in read(book, "xl/worksheets/sheet1.xml")


def test_a_load_onto_occupied_cells_is_refused(loaded: PowerQueryWorkbook) -> None:
    """The fixture already has a table at A1:B3."""
    loaded.add_query("Another", "let Source = 1 in Source")
    with pytest.raises(PowerQueryError, match="already holds something"):
        loaded.load_to_sheet("Another", ["A", "B"], cell="A1")


def test_a_load_with_no_columns_is_refused(book: PowerQueryWorkbook) -> None:
    book.add_query("Loadable", "1")
    with pytest.raises(PowerQueryError, match="at least one column"):
        book.load_to_sheet("Loadable", [])


def test_two_columns_of_one_name_are_refused(book: PowerQueryWorkbook) -> None:
    book.add_query("Loadable", "1")
    with pytest.raises(PowerQueryError, match="share a name"):
        book.load_to_sheet("Loadable", ["A", "A"])


def test_loading_a_query_that_already_loads_is_refused(loaded: PowerQueryWorkbook) -> None:
    with pytest.raises(PowerQueryError, match="already loads"):
        loaded.load_to_sheet("Loaded", ["A"], cell="H1")


def test_a_sheet_that_is_not_there_is_refused(book: PowerQueryWorkbook) -> None:
    book.add_query("Loadable", "1")
    with pytest.raises(PowerQueryError, match="no sheet named"):
        book.load_to_sheet("Loadable", ["A"], sheet="Nowhere")
    with pytest.raises(PowerQueryError, match="not one of them"):
        book.load_to_sheet("Loadable", ["A"], sheet=9)


def test_a_load_can_pick_a_sheet_by_name(book: PowerQueryWorkbook) -> None:
    book.add_query("Loadable", "1")
    book.load_to_sheet("Loadable", ["A"], sheet="Sheet1", cell="H1")
    assert "<tableParts" in read(book, "xl/worksheets/sheet1.xml")


# --- unloading ----------------------------------------------------------------


def test_unloading_takes_every_object_back_out(loaded: PowerQueryWorkbook) -> None:
    assert loaded.unload("Loaded") is True
    names = parts_of(loaded)
    assert "xl/tables/table1.xml" not in names
    assert "xl/queryTables/queryTable1.xml" not in names
    assert "xl/tables/_rels/table1.xml.rels" not in names
    assert "xl/connections.xml" not in names, "an empty connections part is one Excel refuses"
    sheet = read(loaded, "xl/worksheets/sheet1.xml")
    assert "tablePart" not in sheet
    assert "<c r=\"A1\"" not in sheet
    assert "ExternalData" not in read(loaded, "xl/workbook.xml")
    assert "/xl/connections.xml" not in read(loaded, "[Content_Types].xml")


def test_unloading_says_so_in_the_metadata(loaded: PowerQueryWorkbook) -> None:
    loaded.unload("Loaded")
    query = loaded.query("Loaded")
    assert query.load_target == LOAD_CONNECTION_ONLY
    assert query.load_enabled is False
    assert query.target_name is None


def test_unloading_a_query_that_loads_nowhere_says_so(book: PowerQueryWorkbook) -> None:
    assert book.unload("Numbers") is False
    assert book.query("Numbers").load_target == LOAD_CONNECTION_ONLY


def test_a_query_can_be_unloaded_and_loaded_again(loaded: PowerQueryWorkbook) -> None:
    loaded.unload("Loaded")
    loaded.load_to_sheet("Loaded", ["A", "B"], cell="D1")
    assert loaded.query("Loaded").load_target == LOAD_TABLE
    assert 'ref="D1:E2"' in read(loaded, "xl/tables/table1.xml")
    for name in parts_of(loaded):
        if name.endswith((".xml", ".rels")):
            ElementTree.fromstring(read(loaded, name))


def test_the_query_itself_can_load_and_unload(book: PowerQueryWorkbook) -> None:
    book.add_query("Loadable", "let Source = 1 in Source")
    query = book.query("Loadable")
    assert query.load_to_sheet(["A"], cell="H1") == "Loadable"
    assert book.query("Loadable").load_target == LOAD_TABLE
    assert query.unload() is True
    assert book.query("Loadable").load_target == LOAD_CONNECTION_ONLY


def test_a_loaded_workbook_still_saves_and_reopens(book: PowerQueryWorkbook, tmp_path: Path) -> None:
    book.add_query("Loadable", "let Source = 1 in Source")
    book.load_to_sheet("Loadable", ["A", "B"], cell="H1")
    out = book.save(tmp_path / "saved.xlsx")
    again = PowerQueryWorkbook(out)
    assert again.query("Loadable").load_target == LOAD_TABLE
    assert OpcFile.parse(out.read_bytes()).has("xl/tables/table1.xml")


# --- workbooks another writer produced ----------------------------------------
# Excel is not the only thing that writes an .xlsx, and the parts it
# writes are one legal spelling among several.  Each shape below is what
# openpyxl produces, and each defeated an assumption here.


def foreign(tmp_path: Path) -> Path:
    """A minimal workbook spelled the way openpyxl spells one: the
    relationship id written last, the target absolute, an empty
    `definedNames` closed on itself, and no `r:` prefix declared on the
    worksheet, which has no table to need one."""
    out = tmp_path / "foreign.xlsx"
    package = OpcFile()
    package.write(
        "[Content_Types].xml",
        b'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        b'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        b"</Types>",
    )
    package.write(
        "_rels/.rels",
        b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        b' Target="xl/workbook.xml" Id="rId1"/></Relationships>',
    )
    package.write(
        "xl/workbook.xml",
        b'<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        b'<sheets><sheet name="Sheet" sheetId="1" state="visible" r:id="rId1" /></sheets>'
        b"<definedNames /><calcPr calcId=\"124519\" /></workbook>",
    )
    package.write(
        "xl/_rels/workbook.xml.rels",
        b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
        b' Target="/xl/worksheets/sheet1.xml" Id="rId1"/></Relationships>',
    )
    package.write(
        "xl/worksheets/sheet1.xml",
        b'<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<dimension ref="A1:A1"/><sheetData><row r="1">'
        b'<c r="A1" t="inlineStr"><is><t>hi</t></is></c></row></sheetData></worksheet>',
    )
    out.write_bytes(package.serialize())
    return out


def test_a_relationship_is_found_whatever_order_it_is_written_in(tmp_path: Path) -> None:
    """openpyxl closes a relationship with its id where Excel opens with
    it, and XML gives attribute order no meaning.  Matching in a fixed
    order found nothing, and the sheet looked as though it had no part."""
    book = PowerQueryWorkbook(foreign(tmp_path))
    book.add_query("Loaded", 'let\r\n    Source = 1\r\nin\r\n    Source')

    book.load_to_sheet("Loaded", ["N"], cell="C1")

    assert book._opc.has("xl/tables/table1.xml")  # pyright: ignore[reportPrivateUsage]


def test_the_worksheet_gets_the_prefix_its_table_reference_needs(tmp_path: Path) -> None:
    """A worksheet with no table has no use for the `r:` prefix, so
    openpyxl does not declare it.  Adding a `tablePart` that uses it left
    the part not well formed, and Excel would not open the workbook."""
    book = PowerQueryWorkbook(foreign(tmp_path))
    book.add_query("Loaded", 'let\r\n    Source = 1\r\nin\r\n    Source')
    book.load_to_sheet("Loaded", ["N"], cell="C1")

    sheet = book._opc.read("xl/worksheets/sheet1.xml")  # pyright: ignore[reportPrivateUsage]
    assert b"xmlns:r=" in sheet
    ElementTree.fromstring(sheet)


def test_an_empty_defined_names_element_is_filled_not_doubled(tmp_path: Path) -> None:
    """`<definedNames />` is the same element as `<definedNames></...>`.
    Appending a second block beside it puts two in the workbook, which
    Excel refuses."""
    book = PowerQueryWorkbook(foreign(tmp_path))
    book.add_query("Loaded", 'let\r\n    Source = 1\r\nin\r\n    Source')
    book.load_to_sheet("Loaded", ["N"], cell="C1")

    workbook = book._opc.read("xl/workbook.xml").decode("utf-8")  # pyright: ignore[reportPrivateUsage]
    assert workbook.count("<definedNames") == 1
    assert "ExternalData_1" in workbook
    ElementTree.fromstring(workbook)


def test_a_foreign_workbook_round_trips_through_load_and_unload(tmp_path: Path) -> None:
    path = foreign(tmp_path)
    book = PowerQueryWorkbook(path)
    book.add_query("Loaded", 'let\r\n    Source = 1\r\nin\r\n    Source')
    book.load_to_sheet("Loaded", ["N"], cell="C1")
    book.save()

    again = PowerQueryWorkbook(path)
    assert again.query("Loaded").unload() is True
    again.save()
    assert not [n for n in PowerQueryWorkbook(path)._opc.names() if "tables/" in n]  # pyright: ignore[reportPrivateUsage]


def test_openpyxl_itself_if_it_is_installed(tmp_path: Path) -> None:
    """The same path against the real writer, so the shapes above stay
    the ones openpyxl actually produces."""
    openpyxl = pytest.importorskip("openpyxl")

    out = tmp_path / "openpyxl.xlsx"
    made = openpyxl.Workbook()
    made["Sheet"]["A1"] = "hi"
    made.save(out)

    book = PowerQueryWorkbook(out)
    book.add_query("Loaded", 'let\r\n    Source = 1\r\nin\r\n    Source')
    book.load_to_sheet("Loaded", ["N"], cell="C1")
    book.save()

    package = OpcFile.parse(out.read_bytes())
    for part in ("xl/workbook.xml", "xl/worksheets/sheet1.xml"):
        ElementTree.fromstring(package.read(part))
    assert package.read("xl/workbook.xml").decode("utf-8").count("<definedNames") == 1
    assert b"xmlns:r=" in package.read("xl/worksheets/sheet1.xml")
    assert PowerQueryWorkbook(out).query_names() == ["Loaded"]


# --- loading onto a sheet other than the first --------------------------------
# `localSheetId` on a defined name is the zero-based position of the
# sheet the name belongs to.  Writing a constant 0 was right only while
# the table landed on the first sheet; anywhere else the name claimed one
# sheet while its reference named another, and Excel would not open the
# workbook.


def many_sheets(tmp_path: Path, count: int = 3) -> Path:
    openpyxl = pytest.importorskip("openpyxl")

    out = tmp_path / "sheets.xlsx"
    made = openpyxl.Workbook()
    made["Sheet"]["A1"] = "first"
    for index in range(2, count + 1):
        made.create_sheet(f"Sheet{index}")["A1"] = f"sheet {index}"
    made.save(out)
    return out


def defined_names(book: PowerQueryWorkbook) -> list[str]:
    workbook = book._opc.read("xl/workbook.xml").decode("utf-8")  # pyright: ignore[reportPrivateUsage]
    return re.findall(r"<definedName\b[^>]*>[^<]*</definedName>", workbook)


@pytest.mark.parametrize("position", [1, 2, 3])
def test_the_defined_name_names_the_sheet_the_table_is_on(
    tmp_path: Path, position: int
) -> None:
    book = PowerQueryWorkbook(many_sheets(tmp_path))
    book.add_query("Loaded", "let\r\n    Source = 1\r\nin\r\n    Source")
    book.load_to_sheet("Loaded", ["N"], sheet=position, cell="C1")

    name = defined_names(book)[0]
    assert f'localSheetId="{position - 1}"' in name
    sheet = "Sheet" if position == 1 else f"Sheet{position}"
    assert f"{sheet}!$C$1" in name


def test_two_sheets_get_two_names_each_pointing_at_its_own(tmp_path: Path) -> None:
    book = PowerQueryWorkbook(many_sheets(tmp_path))
    for name, position in (("First", 1), ("Third", 3)):
        book.add_query(name, "let\r\n    Source = 1\r\nin\r\n    Source")
        book.load_to_sheet(name, ["N"], sheet=position, cell="C1")

    written = defined_names(book)
    assert len(written) == 2
    assert [re.search(r'localSheetId="(\d+)"', entry).group(1) for entry in written] == ["2", "0"]  # pyright: ignore[reportOptionalMemberAccess]


def test_unloading_from_a_later_sheet_takes_its_name_too(tmp_path: Path) -> None:
    """The removal matches on the name rather than the sheet id, so it has
    to work wherever the table landed."""
    path = many_sheets(tmp_path)
    book = PowerQueryWorkbook(path)
    book.add_query("Loaded", "let\r\n    Source = 1\r\nin\r\n    Source")
    book.load_to_sheet("Loaded", ["N"], sheet=3, cell="C1")
    book.save()

    again = PowerQueryWorkbook(path)
    assert again.query("Loaded").unload() is True
    again.save()

    assert defined_names(PowerQueryWorkbook(path)) == []
    assert not [n for n in PowerQueryWorkbook(path)._opc.names() if "tables/" in n]  # pyright: ignore[reportPrivateUsage]
