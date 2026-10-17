"""Recreate the native eleven-function metadata fixture without instrumentation."""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from pyvbaharness import ExcelSession


def main() -> None:
    folder = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivots'
    baseline = json.loads((folder / 'record_types.json').read_text(encoding='utf-8'))['probes'][0]['source']
    prefix = baseline[:baseline.index('ws.Range("B2").Value=')]
    path = folder / 'native-functions.xlsx'
    saved = folder / 'functions.json'
    record = json.loads(saved.read_text(encoding='utf-8'))
    source = prefix + '''Set c=ThisWorkbook.PivotCaches.Create(xlDatabase,ws.Range("A1:B6"))
Dim functions As Variant, i As Long
functions=Array(xlSum,xlCount,xlCountNums,xlAverage,xlMin,xlMax,xlProduct,xlStDev,xlStDevP,xlVar,xlVarP)
For i=0 To UBound(functions)
Set p=c.CreatePivotTable(ws.Cells(3,5+i*5),"Function" & CStr(i+1))
Set f=p.AddDataField(p.PivotFields("Amount"),,functions(i))
Probe=Probe & p.Name & ":" & CStr(functions(i)) & ":" & CStr(f.Function) & ";"
Next
ThisWorkbook.SaveAs "''' + str(path) + '''",xlOpenXMLWorkbook
Exit Function
Failed:
Probe="ERR:" & CStr(Err.Number)
End Function'''
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\n'
                               'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                               'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        result = excel.run_vba(source, 'Probe', line_numbers=False)
        assert result.ok and not result.dialogs and result.value == record['creation_result'], result
        observations: list[str] = []
        for _ in range(3):
            excel.open_document(path)
            inspected = excel.run_vba(record['inspection_source'], 'InspectFunctions', line_numbers=False)
            assert inspected.ok and not inspected.dialogs and isinstance(inspected.value, str), inspected
            assert sorted(inspected.value.split(';')) == sorted(record['inspection_result'].split(';'))
            observations.append(inspected.value)
        record['inspection_result'] = observations[-1]
        record['repeated_open_observations'] = observations
        record['excel_version_build'] = version.value
    record['measured_at'] = dt.datetime.now(dt.timezone.utc).isoformat()
    saved.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
