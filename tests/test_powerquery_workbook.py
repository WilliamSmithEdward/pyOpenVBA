"""Reading and writing the queries of a workbook.

The fixtures were built by Excel itself, so what the reader says about
them is checked against what Excel put there.  The live gate in
``tests/test_live_powerquery_gate.py`` takes the other half: it has Excel
open what this writes and evaluate it.
"""

from __future__ import annotations

import shutil
import uuid
from pathlib import Path

import pytest

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery import _metadata as meta
from pyopenvba.powerquery import LOAD_CONNECTION_ONLY, LOAD_TABLE, PowerQueryWorkbook

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"
WORKBOOKS = sorted(FIXTURES.glob("*.xlsx"))


@pytest.fixture
def book(tmp_path: Path) -> PowerQueryWorkbook:
    out = tmp_path / "three.xlsx"
    shutil.copyfile(FIXTURES / "three_queries.xlsx", out)
    return PowerQueryWorkbook(out)


@pytest.fixture
def empty(tmp_path: Path) -> PowerQueryWorkbook:
    out = tmp_path / "none.xlsx"
    shutil.copyfile(FIXTURES / "no_queries.xlsx", out)
    return PowerQueryWorkbook(out)


# --- reading ------------------------------------------------------------------


def test_a_workbook_lists_the_queries_excel_put_in_it(book: PowerQueryWorkbook) -> None:
    assert book.query_names() == ["Numbers", "Doubled", "Count Of Rows"]
    assert book.has_queries


def test_a_workbook_without_a_package_has_no_queries(empty: PowerQueryWorkbook) -> None:
    assert empty.query_names() == []
    assert not empty.has_queries
    assert empty.groups() == []


def test_a_query_carries_its_m_and_what_excel_recorded(book: PowerQueryWorkbook) -> None:
    numbers = book.query("Numbers")
    assert numbers.formula.startswith("let")
    assert "{1..10}" in numbers.formula
    assert numbers.steps == ["Source", "Table"]
    assert numbers.load_target == LOAD_CONNECTION_ONLY
    assert numbers.load_enabled is False
    assert numbers.is_private is False
    assert uuid.UUID(str(numbers.query_id))
    assert numbers.description is None
    assert numbers.is_function is False


def test_a_loaded_query_says_where_it_loads() -> None:
    book = PowerQueryWorkbook(FIXTURES / "loaded_to_sheet.xlsx")
    loaded = book.query("Loaded")
    assert loaded.load_target == LOAD_TABLE
    assert loaded.load_enabled is True
    assert loaded.target_name == "Loaded"


def test_a_described_query_gives_back_its_description() -> None:
    book = PowerQueryWorkbook(FIXTURES / "described.xlsx")
    assert book.query("Described").description == "A described query"


def test_names_that_need_quoting_read_as_themselves() -> None:
    book = PowerQueryWorkbook(FIXTURES / "odd_names.xlsx")
    names = book.query_names()
    assert "With Space" in names
    assert "Café" in names
    assert "日本" in names
    assert "let" in names


def test_a_function_query_is_marked_as_one() -> None:
    book = PowerQueryWorkbook(FIXTURES / "functions.xlsx")
    assert book.query("Arrow").is_function
    assert book.query("EachFn").is_function
    assert not book.query("Rec").is_function
    assert str(book.query("Arrow").entries()[meta.RESULT_TYPE]) == "Function"


def test_the_groups_excel_kept_read_back() -> None:
    """This fixture was written here and then edited by Excel, which
    preserved the groups and the query's place in one."""
    book = PowerQueryWorkbook(FIXTURES / "grouped.xlsx")
    assert [group.name for group in book.groups()] == ["Staging"]
    assert book.query("Numbers").group is not None
    assert book.query("Numbers").group.name == "Staging"  # pyright: ignore[reportOptionalMemberAccess]
    assert book.query("Doubled").group is None


def test_asking_for_a_query_that_is_not_there_is_refused(book: PowerQueryWorkbook) -> None:
    with pytest.raises(PowerQueryError, match="no query named"):
        book.query("Nothing")


def test_a_file_that_is_not_an_excel_package_is_refused(tmp_path: Path) -> None:
    path = tmp_path / "notes.txt"
    path.write_text("hello", encoding="utf-8")
    with pytest.raises(PowerQueryError, match="not an Excel package"):
        PowerQueryWorkbook(path)


# --- writing ------------------------------------------------------------------


@pytest.mark.parametrize("path", WORKBOOKS, ids=lambda p: p.name)
def test_a_workbook_that_was_not_changed_is_written_back_byte_for_byte(path: Path) -> None:
    assert PowerQueryWorkbook(path).to_bytes() == path.read_bytes()


def test_a_formula_change_reaches_the_section_and_the_steps(book: PowerQueryWorkbook) -> None:
    book.query("Numbers").formula = "let A = 1, B = A in B"
    assert book.query("Numbers").formula == "let A = 1, B = A in B"
    assert book.query("Numbers").steps == ["A", "B"]
    assert [item.parts[2] for item in book.metadata.steps_of("Numbers")] == ["A", "B"]
    saved = PowerQueryWorkbook(book.save(book.path))
    assert saved.query("Numbers").formula == "let A = 1, B = A in B"


def test_an_added_query_gets_the_entries_excel_gives_a_new_one(book: PowerQueryWorkbook) -> None:
    added = book.add_query("Extra", "let\r\n    Source = 99\r\nin\r\n    Source")
    assert added.name == "Extra"
    assert book.query_names() == ["Numbers", "Doubled", "Count Of Rows", "Extra"]
    entries = added.entries()
    assert entries[meta.IS_PRIVATE] == 0
    assert entries[meta.FILL_ENABLED] == 0
    assert entries[meta.FILL_OBJECT_TYPE] == "ConnectionOnly"
    assert entries[meta.FILL_TO_DATA_MODEL_ENABLED] == 0
    assert uuid.UUID(str(entries[meta.QUERY_ID]))
    assert [item.parts[2] for item in book.metadata.steps_of("Extra")] == ["Source"]


def test_an_added_function_is_marked_as_one(book: PowerQueryWorkbook) -> None:
    added = book.add_query("Add", "(x as number, y as number) as number => x + y")
    assert added.entries()[meta.RESULT_TYPE] == "Function"
    assert book.metadata.steps_of("Add") == []


def test_an_added_query_can_carry_a_description(book: PowerQueryWorkbook) -> None:
    book.add_query("Extra", "1", description="a note")
    assert book.query("Extra").description == "a note"


def test_adding_a_name_already_there_is_refused(book: PowerQueryWorkbook) -> None:
    with pytest.raises(PowerQueryError, match="already has a query"):
        book.add_query("Numbers", "1")


def test_adding_a_query_with_no_name_is_refused(book: PowerQueryWorkbook) -> None:
    with pytest.raises(PowerQueryError, match="needs a name"):
        book.add_query("", "1")


def test_a_removed_query_takes_its_item_and_steps_with_it(book: PowerQueryWorkbook) -> None:
    book.remove_query("Count Of Rows")
    assert book.query_names() == ["Numbers", "Doubled"]
    assert book.metadata.query("Count Of Rows") is None
    assert book.metadata.steps_of("Count Of Rows") == []
    assert "Count Of Rows" not in book.section_text()


def test_a_rename_moves_the_member_the_item_and_the_references(book: PowerQueryWorkbook) -> None:
    book.rename_query("Numbers", "Renamed Numbers")
    assert book.query_names() == ["Renamed Numbers", "Doubled", "Count Of Rows"]
    assert book.metadata.query("Renamed Numbers") is not None
    assert [item.parts[2] for item in book.metadata.steps_of("Renamed Numbers")] == ["Source", "Table"]
    assert '#"Renamed Numbers"' in book.query("Doubled").formula


def test_a_rename_can_leave_the_references_alone(book: PowerQueryWorkbook) -> None:
    book.rename_query("Numbers", "Renamed", update_references=False)
    assert "Numbers" in book.query("Doubled").formula


def test_a_rename_onto_a_name_already_there_is_refused(book: PowerQueryWorkbook) -> None:
    with pytest.raises(PowerQueryError, match="already has a query"):
        book.rename_query("Numbers", "Doubled")


def test_a_description_can_be_set_and_dropped(book: PowerQueryWorkbook) -> None:
    query = book.query("Numbers")
    query.description = "one to ten"
    assert book.query("Numbers").description == "one to ten"
    query.description = None
    assert book.query("Numbers").description is None


def test_a_group_can_be_added_and_a_query_moved_into_it(book: PowerQueryWorkbook) -> None:
    staging = book.add_group("Staging")
    nested = book.add_group("Raw", parent=staging)
    book.query("Numbers").move_to_group(nested)
    assert [group.name for group in book.groups()] == ["Staging", "Raw"]
    assert book.query("Numbers").group is not None
    assert book.query("Numbers").group.name == "Raw"  # pyright: ignore[reportOptionalMemberAccess]
    saved = PowerQueryWorkbook(book.save(book.path))
    assert saved.query("Numbers").group.name == "Raw"  # pyright: ignore[reportOptionalMemberAccess]


def test_a_group_removed_leaves_its_queries_at_the_top(book: PowerQueryWorkbook) -> None:
    staging = book.add_group("Staging")
    book.query("Numbers").move_to_group(staging)
    book.remove_group(staging)
    assert book.groups() == []
    assert book.query("Numbers").group is None


def test_two_groups_of_the_same_name_in_one_place_are_refused(book: PowerQueryWorkbook) -> None:
    book.add_group("Staging")
    with pytest.raises(PowerQueryError, match="already has a group"):
        book.add_group("Staging")


def test_a_workbook_with_no_package_gets_one(empty: PowerQueryWorkbook) -> None:
    """The hardest write: the custom XML part, its properties, the
    relationships and the content type all have to appear."""
    empty.add_query("FromNothing", "let\r\n    Source = 1\r\nin\r\n    Source")
    out = empty.save()
    reopened = PowerQueryWorkbook(out)
    assert reopened.query_names() == ["FromNothing"]
    assert reopened.query("FromNothing").formula.startswith("let")
    names = reopened._opc.names()  # pyright: ignore[reportPrivateUsage]
    assert "customXml/item1.xml" in names
    assert "customXml/itemProps1.xml" in names
    assert "customXml/_rels/item1.xml.rels" in names
    types = reopened._opc.read("[Content_Types].xml").decode()  # pyright: ignore[reportPrivateUsage]
    assert "/customXml/itemProps1.xml" in types
    rels = reopened._opc.read("xl/_rels/workbook.xml.rels").decode()  # pyright: ignore[reportPrivateUsage]
    assert "../customXml/item1.xml" in rels


def test_a_saved_workbook_reopens_with_the_same_queries(book: PowerQueryWorkbook, tmp_path: Path) -> None:
    book.add_query("Extra", "let Source = 1 in Source")
    book.remove_query("Doubled")
    out = book.save(tmp_path / "edited.xlsx")
    reopened = PowerQueryWorkbook(out)
    assert reopened.query_names() == ["Numbers", "Count Of Rows", "Extra"]


def test_saving_elsewhere_leaves_the_original_alone(book: PowerQueryWorkbook, tmp_path: Path) -> None:
    before = book.path.read_bytes()
    book.add_query("Extra", "1")
    book.save(tmp_path / "copy.xlsx")
    assert book.path.read_bytes() == before


def test_the_whole_section_can_be_replaced(book: PowerQueryWorkbook) -> None:
    book.set_section_text("section Section1;\r\n\r\nshared Only = let A = 1 in A;")
    assert book.query_names() == ["Only"]
    assert book.metadata.query("Numbers") is None
    assert [item.parts[2] for item in book.metadata.steps_of("Only")] == ["A"]
    assert book.query("Only").entries()[meta.FILL_OBJECT_TYPE] == "ConnectionOnly"


def test_the_signature_goes_when_the_package_changes(book: PowerQueryWorkbook) -> None:
    assert book.mashup.bindings
    book.add_query("Extra", "1")
    assert book.mashup.bindings == b""
