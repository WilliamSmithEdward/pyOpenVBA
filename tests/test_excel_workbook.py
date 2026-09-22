"""The in-memory workbook: what a macro does to it, and what a save keeps.

The object model's answers are held to live Excel in
``test_excel_model.py``.  This covers the parts around it: loading a
real file, writing one back, and what the Python side of the API hands
back.
"""

from __future__ import annotations

import datetime as _dt
import shutil
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.powerquery._opc import OpcFile

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"
LOADED = FIXTURES / "loaded_to_sheet.xlsx"
THREE_QUERIES = FIXTURES / "three_queries.xlsx"


def fresh() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    return app


def macro(app: ExcelApplication, body: str, name: str = "Main") -> object:
    app.add_module(f"Sub {name}()\n{body}\nEnd Sub\n", name=f"Module_{name}")
    return app.run(name)


# --- state from nothing ---------------------------------------------------------------


def test_a_new_workbook_has_one_sheet() -> None:
    app = fresh()
    assert [sheet.name for sheet in app.sheets()] == ["Sheet1"]


def test_a_macro_writes_cells_that_python_reads_back() -> None:
    app = fresh()
    macro(app, '    Range("A1").Value = 5\n    Range("B1").Value = "five"')
    assert app.sheet(1).value("A1") == 5
    assert app.sheet(1).value("B1") == "five"


def test_a_macro_adds_a_sheet_before_the_active_one() -> None:
    """Add with no arguments puts the sheet in front of the active one."""
    app = fresh()
    macro(app, '    Worksheets.Add.Name = "Second"')
    assert [sheet.name for sheet in app.sheets()] == ["Second", "Sheet1"]


def test_add_after_puts_the_sheet_where_it_was_asked() -> None:
    app = fresh()
    macro(app, '    Worksheets.Add(After:=Worksheets(1)).Name = "Second"')
    assert [sheet.name for sheet in app.sheets()] == ["Sheet1", "Second"]


def test_two_sheets_cannot_share_a_name() -> None:
    app = fresh()
    with pytest.raises(VBARuntimeError) as raised:
        macro(app, '    Worksheets.Add.Name = "Sheet1"')
    assert raised.value.number == 1004


def test_a_block_written_from_an_array_lands_row_by_row() -> None:
    app = fresh()
    app.add_module(
        "Sub Main()\n"
        "    Dim grid(1 To 2, 1 To 2) As Variant\n"
        "    grid(1, 1) = 1\n"
        "    grid(1, 2) = 2\n"
        "    grid(2, 1) = 3\n"
        "    grid(2, 2) = 4\n"
        '    Range("A1:B2").Value = grid\n'
        "End Sub\n",
        name="Module1",
    )
    app.run("Main")
    assert app.sheet(1).rows() == [[1, 2], [3, 4]]


def test_a_block_read_back_is_a_two_dimensional_array() -> None:
    app = fresh()
    macro(app, '    Range("A1").Value = 1\n    Range("B2").Value = 4')
    assert app.evaluate('Range("A1:B2").Value') == [[1, None], [None, 4]]


def test_the_console_holds_what_debug_print_wrote() -> None:
    app = fresh()
    macro(app, '    Debug.Print "a", "b"')
    assert app.console == ["a             b"]


def test_describe_shows_the_grid() -> None:
    app = fresh()
    macro(app, '    Range("A1").Value = "x"\n    Range("B2").Value = 2')
    dump = app.describe()
    assert "Worksheet 'Sheet1' (A1:B2)" in dump
    assert "x" in dump


# --- the formula gap is reported, not guessed -------------------------------------------


def test_a_formula_written_by_a_macro_is_kept_as_text() -> None:
    app = fresh()
    macro(app, '    Range("A1").Value = 1\n    Range("A2").Formula = "=A1+1"')
    assert app.sheet(1).formula("A2") == "=A1+1"


def test_a_formula_written_by_a_macro_is_calculated_when_it_is_read() -> None:
    app = fresh()
    macro(app, '    Range("A1").Value = 2\n    Range("A2").Formula = "=A1*3"')
    assert app.sheet(1).value("A2") == 6.0


def test_a_change_upstream_reaches_the_formula_that_reads_it() -> None:
    app = fresh()
    macro(app, '    Range("A1").Value = 2\n    Range("A2").Formula = "=A1*3"')
    assert app.sheet(1).value("A2") == 6.0
    macro(app, '    Range("A1").Value = 5', name="Again")
    assert app.sheet(1).value("A2") == 15.0


def test_manual_calculation_keeps_the_value_until_calculate_is_called() -> None:
    """Manual mode is what Excel does until F9: the old value stands."""
    app = fresh()
    macro(
        app,
        '    Range("A1").Value = 2\n'
        '    Range("A2").Formula = "=A1*3"\n'
        '    Debug.Print Range("A2").Value\n'
        "    Application.Calculation = xlCalculationManual\n"
        '    Range("A1").Value = 100',
    )
    assert app.sheet(1).value("A2") == 6.0
    macro(app, "    Application.Calculate", name="Recalc")
    assert app.sheet(1).value("A2") == 300.0


# --- unknown members are answered honestly ------------------------------------------------


def test_a_real_member_this_lacks_is_unsupported() -> None:
    app = fresh()
    with pytest.raises(VBAUnsupportedError) as raised:
        macro(app, "    Worksheets(1).PivotTables.Count")
    assert "PivotTables" in str(raised.value)


def test_a_member_excel_has_not_got_either_is_error_438() -> None:
    app = fresh()
    with pytest.raises(VBARuntimeError) as raised:
        macro(app, "    Worksheets(1).Pivottabel")
    assert raised.value.number == 438


def test_a_worksheet_function_this_lacks_says_which_one() -> None:
    app = fresh()
    with pytest.raises(VBAUnsupportedError) as raised:
        macro(app, "    Application.WorksheetFunction.Sumproduct(1, 2)")
    assert "Sumproduct" in str(raised.value)


# --- loading a real workbook ----------------------------------------------------------------


def test_opening_a_workbook_reads_its_cells() -> None:
    app = ExcelApplication.open(LOADED)
    assert app.sheet(1).rows() == [["A", "B"], [1, "x"], [2, "y"]]


def test_opening_a_workbook_reads_its_names_and_queries() -> None:
    app = ExcelApplication.open(LOADED)
    book = app.workbook
    assert [entry.name for entry in book.names_.entries] == ["Sheet1!ExternalData_1"]
    assert [entry.name for entry in book.queries_.entries] == ["Loaded"]


def test_a_query_refresh_evaluates_its_m() -> None:
    """Refresh runs the query rather than reporting itself unsupported."""
    app = ExcelApplication.open(LOADED)
    macro(app, '    ActiveWorkbook.Queries("Loaded").Refresh')
    assert app.sheet(1).rows() == [["A", "B"], [1, "x"], [2, "y"]]


def test_a_query_formula_can_be_read_through_vba() -> None:
    app = ExcelApplication.open(THREE_QUERIES)
    names = app.evaluate("ActiveWorkbook.Queries.Count")
    assert names == 3


# --- writing one back --------------------------------------------------------------------------


def test_a_workbook_nobody_changed_saves_byte_for_byte(tmp_path: Path) -> None:
    """Nothing is rewritten unless the model changed it.

    A save that regenerated every sheet would quietly drop whatever the
    model has no field for, which for this fixture is the whole Power
    Query load: its connection, its table and its query table.
    """
    copy = tmp_path / "same.xlsx"
    shutil.copyfile(LOADED, copy)
    app = ExcelApplication.open(copy)
    out = tmp_path / "out.xlsx"
    app.save(out)
    assert out.read_bytes() == LOADED.read_bytes()


def test_a_saved_change_keeps_every_other_part(tmp_path: Path) -> None:
    copy = tmp_path / "changed.xlsx"
    shutil.copyfile(LOADED, copy)
    app = ExcelApplication.open(copy)
    macro(app, '    Worksheets(1).Range("D1").Value = "added"')
    out = tmp_path / "out.xlsx"
    app.save(out)

    before = OpcFile.parse(LOADED.read_bytes())
    after = OpcFile.parse(out.read_bytes())
    assert after.names() == before.names()
    changed = [name for name in before.names() if before.read(name) != after.read(name)]
    assert changed == ["xl/worksheets/sheet1.xml"]


def test_what_was_written_reads_back_the_same(tmp_path: Path) -> None:
    copy = tmp_path / "roundtrip.xlsx"
    shutil.copyfile(LOADED, copy)
    app = ExcelApplication.open(copy)
    app.add_module(
        "Sub Main()\n"
        '    Worksheets(1).Range("D1").Value = "text"\n'
        '    Worksheets(1).Range("D2").Value = 1234.5\n'
        '    Worksheets(1).Range("D3").Value = True\n'
        '    Worksheets(1).Range("D4").Value = #3/4/2021#\n'
        '    Worksheets(1).Range("D4").NumberFormat = "yyyy-mm-dd"\n'
        "End Sub\n",
        name="Module1",
    )
    app.run("Main")
    out = tmp_path / "out.xlsx"
    app.save(out)

    again = ExcelApplication.open(out)
    view = again.sheet(1)
    assert view.value("D1") == "text"
    assert view.value("D2") == 1234.5
    assert view.value("D3") is True
    assert view.value("D4") == _dt.datetime(2021, 3, 4)
    assert again.evaluate('Worksheets(1).Range("D4").Text') == "2021-03-04"


def test_a_workbook_built_from_nothing_saves_and_reopens(tmp_path: Path) -> None:
    app = fresh()
    app.add_module(
        "Sub Main()\n"
        '    Worksheets(1).Name = "Made"\n'
        '    Range("A1").Value = "made here"\n'
        "    Range(\"A2\").Value = 7\n"
        "End Sub\n",
        name="Module1",
    )
    app.run("Main")
    out = tmp_path / "made.xlsx"
    app.save(out)

    again = ExcelApplication.open(out)
    assert again.sheet(1).name == "Made"
    assert again.sheet(1).rows() == [["made here"], [7]]


def test_a_new_workbook_saved_as_xlsx_has_no_macro_part(tmp_path: Path) -> None:
    app = fresh()
    macro(app, '    Range("A1").Value = 1')
    out = tmp_path / "plain.xlsx"
    app.save(out)
    assert "xl/vbaProject.bin" not in OpcFile.parse(out.read_bytes()).names()


def test_a_cleared_cell_leaves_the_file(tmp_path: Path) -> None:
    copy = tmp_path / "cleared.xlsx"
    shutil.copyfile(LOADED, copy)
    app = ExcelApplication.open(copy)
    macro(app, '    Worksheets(1).Range("B2").ClearContents')
    out = tmp_path / "out.xlsx"
    app.save(out)
    assert ExcelApplication.open(out).sheet(1).value("B2") is None


# --- the clock and the dialogs ----------------------------------------------------------------


def test_a_frozen_clock_makes_now_repeatable() -> None:
    app = fresh()
    app.freeze_clock(_dt.datetime(2021, 3, 4, 12, 30))
    macro(app, '    Range("A1").Value = Now')
    assert app.sheet(1).value("A1") == _dt.datetime(2021, 3, 4, 12, 30)


def test_a_queued_answer_reaches_the_macro() -> None:
    app = fresh()
    app.answer_dialogs(6)
    macro(app, '    If MsgBox("ok?", vbYesNo) = vbYes Then Range("A1").Value = "yes"')
    assert app.sheet(1).value("A1") == "yes"
    assert app.dialogs[0][1] == "ok?"
