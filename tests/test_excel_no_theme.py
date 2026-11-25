"""A workbook with no theme part takes the theme a new workbook has now for its colours: colours its cells hold as theme
colours and ones a macro sets resolve with today's accent 1, 156082 (tests/fixtures/no_theme.json). The built-in styles
Excel adds to such a book take the 2007 theme instead, which tests/test_excel_tool_book_spellings.py holds."""
from __future__ import annotations

import base64
import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/no_theme.json").read_text(encoding="utf-8"))


def test_colours_resolve_with_todays_theme(tmp_path: Path) -> None:
    book = tmp_path / "no_theme.xlsx"
    book.write_bytes(base64.b64decode(RECORD["book"]))
    reads = " & \",\" & ".join(RECORD["reads"])
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f"""Public Function Probe(ByVal path As String) As String
Dim wb, ws
Set wb = Workbooks.Open(path): Set ws = wb.Worksheets(1)
{RECORD["setup"]}
Probe = {reads}
wb.Close False
End Function
""", name="M")
    assert str(app.run("M.Probe", str(book))).split(",") == RECORD["excel"]
