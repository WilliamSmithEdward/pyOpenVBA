"""Excel opens the model's format conversions without repair and identifies them correctly."""

from __future__ import annotations

import importlib
import os
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication


@pytest.mark.skipif(os.environ.get("RUN_LIVE_EXCEL") != "1", reason="requires isolated real Excel")
def test_native_excel_reads_save_as_formats(tmp_path: Path) -> None:
    harness = importlib.import_module("pyvbaharness")

    app = ExcelApplication.open(Path(__file__).parent / "fixtures/shapes/excel_shapes.xlsm", with_vba=False)
    app.workbook.SaveAs(str(tmp_path / "macro-free.xlsx"), 51)
    app.workbook.SaveAs(str(tmp_path / "converted.xlsm"), 52)
    fresh = ExcelApplication()
    fresh.add_workbook()
    fresh.sheet(1).set_value("A1", 42)
    fresh.workbook.SaveAs(str(tmp_path / "fresh.xlsm"), 52)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        for name, file_format, project in (("macro-free.xlsx", 51, False), ("converted.xlsm", 52, False), ("fresh.xlsm", 52, True)):
            source = 'Function Probe() As String\nDim book As Workbook\nSet book = Workbooks.Open("' + str(tmp_path / name).replace('"', '""') + '")\nProbe = CStr(book.FileFormat) & ":" & CStr(book.HasVBProject)\nbook.Close False\nEnd Function'
            result = excel.run_vba(source, "Probe", timeout=30.0)
            assert result.ok, result
            assert not result.dialogs
            assert result.value == f"{file_format}:{project}"
