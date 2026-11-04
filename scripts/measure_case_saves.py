"""SaveAs, SaveCopyAs and Open over a file in another case. Rival rules for a workbook's name after SaveAs: as typed (ALPHA.XLSX), or as the disk spells it (Alpha.xlsx).

    python scripts/measure_case_saves.py

writes tests/fixtures/case_saves.json, which tests/test_excel_link_edits.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/case_saves.json'
SOURCE = r'''Public Function Tried(ByVal folder As String) As String
    Dim wb As Workbook, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.SaveAs folder & "\Alpha.xlsx", xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.SaveAs folder & "\ALPHA.XLSX", xlOpenXMLWorkbook
    s = "saveas " & wb.Name & " dir " & Dir(folder & "\*.xlsx")
    wb.Save
    wb.Close False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.SaveCopyAs folder & "\alpha.XLSX"
    s = s & "|copy dir " & Dir(folder & "\*.xlsx")
    wb.Close False
    Set wb = Workbooks.Open(folder & "\ALPHA.xlsx", 0)
    s = s & "|open " & wb.Name
    wb.Close False
    Tried = s
End Function
'''


def main() -> None:
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder, ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        result = excel.run_vba(SOURCE, 'Tried', (folder,), timeout=60.0)
        assert result.ok, result
        files = sorted(path.name for path in Path(folder).iterdir())
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'answer': str(result.value), 'files': files}, indent=1) + '\n', encoding='utf-8')
    print(result.value, files)


if __name__ == '__main__':
    main()
