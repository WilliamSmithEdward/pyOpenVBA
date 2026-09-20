"""Live Excel gate for the in-memory workbook (opt-in).

Two things only real Excel can answer.  First, that a file this wrote
opens without a repair prompt and holds what the macro put in it: the
offline tests prove the bytes are what we meant, not that Excel agrees.
Second, that a macro run here and the same macro run in Excel leave the
same cells behind.

Opt-in: set ``RUN_LIVE_EXCEL=1`` on a Windows machine with desktop Excel
installed.  Skipped everywhere else, including CI.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_EXCEL") != "1" or sys.platform != "win32",
    reason="live Excel gate: set RUN_LIVE_EXCEL=1 on Windows with Excel installed",
)

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"

#: The macro both sides run.  Everything it touches is something the
#: model claims to implement.
MACRO = """
Sub Fill()
    Dim r As Long
    Range("A1").Value = "Item"
    Range("B1").Value = "Qty"
    For r = 1 To 5
        Cells(r + 1, 1).Value = "Item " & r
        Cells(r + 1, 2).Value = r * 10
    Next r
    Range("D1").Value = Application.WorksheetFunction.Sum(Range("B2:B6"))
    Range("D2").Value = Range("A1").End(xlDown).Address
    Range("D3").Value = ActiveSheet.UsedRange.Address
    Range("D4").Value = Range("B2:B6").Count
End Sub
"""


def _excel_says(source: str, macro: str) -> dict[str, object]:
    """Run the macro in real Excel and read the cells back."""
    harness = pytest.importorskip("pyvbaharness")

    reader = (
        "Public Function Report() As String\n"
        "    Fill\n"
        '    Report = Range("D1").Value & "|" & Range("D2").Value & "|" & _\n'
        '        Range("D3").Value & "|" & Range("D4").Value & "|" & _\n'
        '        Range("A6").Value & "|" & TypeName(Range("B2").Value)\n'
        "End Function\n"
    )
    with harness.ExcelSession() as excel:
        excel.new_document()
        result = excel.run_vba(macro + reader, "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        pieces = str(result.value).split("|")
    return dict(zip(("total", "end", "used", "count", "a6", "type"), pieces))


def _model_says(macro: str) -> dict[str, object]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(macro, name="Module1")
    app.run("Fill")
    sheet = app.sheet(1)
    return {
        "total": app.evaluate('CStr(Range("D1").Value)'),
        "end": sheet.value("D2"),
        "used": sheet.value("D3"),
        "count": app.evaluate('CStr(Range("D4").Value)'),
        "a6": sheet.value("A6"),
        "type": app.evaluate('TypeName(Range("B2").Value)'),
    }


def test_the_model_and_excel_agree_on_what_the_macro_did() -> None:
    assert _model_says(MACRO) == _excel_says("", MACRO)


def test_a_workbook_this_wrote_opens_in_excel(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")

    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MACRO, name="Module1")
    app.run("Fill")
    out = tmp_path / "written.xlsx"
    app.save(out)

    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(
            "Public Function Report() As String\n"
            '    Report = ActiveSheet.Range("A1").Value & "|" & _\n'
            '        CStr(ActiveSheet.Range("B6").Value) & "|" & _\n'
            '        ActiveSheet.UsedRange.Address\n'
            "End Function\n",
            "Report",
            timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "Item|50|$A$1:$D$6"


def test_an_edited_workbook_still_opens_with_its_query(tmp_path: Path) -> None:
    """A Power Query workbook edited here keeps what Excel needs.

    The model has no field for a connection or a query table, so this is
    the check that not rewriting them was enough.
    """
    harness = pytest.importorskip("pyvbaharness")

    copy = tmp_path / "edited.xlsx"
    shutil.copyfile(FIXTURES / "loaded_to_sheet.xlsx", copy)
    app = ExcelApplication.open(copy)
    app.add_module('Sub Touch()\n    Worksheets(1).Range("E1").Value = 99\nEnd Sub\n', name="Module1")
    app.run("Touch")
    out = tmp_path / "edited_out.xlsx"
    app.save(out)

    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(
            "Public Function Report() As String\n"
            '    Report = CStr(ActiveSheet.Range("E1").Value) & "|" & _\n'
            "        CStr(ActiveWorkbook.Queries.Count) & \"|\" & _\n"
            "        CStr(ActiveWorkbook.Connections.Count)\n"
            "End Function\n",
            "Report",
            timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "99|1|1"
