"""A module whose procedures carry attributes (Application.MacroOptions wrote them) exported with VBComponent.Export
and imported into another workbook with VBComponents.Import: the exported file, the imported module's CountOfLines
and the module that workbook then saves.

    python scripts/measure_vbe_import_attributes.py

writes tests/fixtures/vbe_import_attributes.json, which tests/test_excel_vbe_import_attributes.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from pyopenvba.excel import ExcelFile

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vbe_import_attributes.json'
SOURCE = '''Public Function Probe(ByVal folder As String) As String
    Dim a As Workbook, b As Workbook, s As String
    Application.DisplayAlerts = False
    Set a = Workbooks.Add(xlWBATWorksheet)
    a.VBProject.VBComponents.Add(1).CodeModule.AddFromString "Function F(x)" & vbCrLf & "F = x" & vbCrLf & _
        "End Function" & vbCrLf & "Sub S()" & vbCrLf & "End Sub"
    a.Activate
    Application.MacroOptions Macro:="F", Description:="d", Category:=5
    Application.MacroOptions Macro:="S", HasShortcutKey:=True, ShortcutKey:="k"
    a.VBProject.VBComponents("Module1").Export folder & "\\m.bas"
    Set b = Workbooks.Add(xlWBATWorksheet)
    b.VBProject.VBComponents.Import folder & "\\m.bas"
    s = b.VBProject.VBComponents("Module1").CodeModule.CountOfLines & "|" & _
        Replace(b.VBProject.VBComponents("Module1").CodeModule.Lines(1, 5), vbCrLf, "|")
    b.VBProject.VBComponents("Module1").CodeModule.InsertLines 1, "Dim top As Long"
    b.SaveAs folder & "\\b.xlsm", xlOpenXMLWorkbookMacroEnabled
    b.Close False
    a.Close False
    Application.DisplayAlerts = True
    Probe = s
End Function
'''


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(SOURCE, 'Probe', (folder,), timeout=200.0)
        assert result.ok and not result.dialogs, result
        exported = (Path(folder) / 'm.bas').read_bytes().decode('latin-1')
        module = ExcelFile(Path(folder) / 'b.xlsm').get_module('Module1')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'answer': result.value,
              'exported': exported, 'module': module}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print(record['answer'])
    print(repr(exported))
    print(repr(module))


if __name__ == '__main__':
    main()
