"""Putting a query's result on a worksheet, and taking it off again.

What the metadata says about where a query loads is not what makes it
load: Excel needs the connection, the query table, the table and the
sheet's reference to it.  These check that all of that is written, and
that unloading takes every piece back out -- including the connections
part itself, which Excel refuses when it holds no connections.
"""

from __future__ import annotations

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
