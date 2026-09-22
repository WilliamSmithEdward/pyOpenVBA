"""Measure worksheet Move ordering, ownership, activation and names."""
import json
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

PROBES = {
    'before': 'src.Worksheets("Third").Move Before:=src.Worksheets(1)',
    'after': 'src.Worksheets(1).Move After:=src.Worksheets("Third")',
    'self': 'src.Worksheets(1).Move Before:=src.Worksheets(1)',
    'new': 'src.Worksheets(1).Move\nSet dst = ActiveWorkbook',
    'cross': 'src.Worksheets(1).Move After:=dst.Worksheets(1)',
    'names': 'src.Names.Add "GlobalAmount", "=Sheet1!$A$1"\nsrc.Worksheets(1).Names.Add "LocalAmount", "=Sheet1!$A$1"\nsrc.Worksheets(1).Range("B1").Formula = "=GlobalAmount+LocalAmount"\nsrc.Worksheets(1).Move Before:=dst.Worksheets(1)',
    'last': 'src.Worksheets("Third").Delete\nsrc.Worksheets("Second").Delete\nsrc.Worksheets(1).Move After:=dst.Worksheets(1)',
    'last_new': 'src.Worksheets("Third").Delete\nsrc.Worksheets("Second").Delete\nOn Error Resume Next\nsrc.Worksheets(1).Move\nReport = CStr(Err.Number)\nExit Function',
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        for name, statement in PROBES.items():
            excel.new_document()
            body = ('Set src = Workbooks.Add\nSet dst = Workbooks.Add\nApplication.DisplayAlerts = False\n'
                    'src.Worksheets.Add(After:=src.Worksheets(1)).Name = "Second"\n'
                    'src.Worksheets.Add(After:=src.Worksheets(2)).Name = "Third"\n'
                    'src.Worksheets(1).Range("A1").Value = 7\n' + statement +
                    '\nReport = CStr(ActiveWorkbook Is dst) & ":" & ActiveSheet.Name & "|"\n'
                    'For Each bk In Workbooks\nIf bk Is src Or bk Is dst Then\nIf bk Is src Then\nReport = Report & "source:"\nElse\nReport = Report & "destination:"\nEnd If\n'
                    'For Each sh In bk.Worksheets\nReport = Report & sh.Name & ","\nNext sh\n'
                    'For Each nm In bk.Names\nReport = Report & nm.Name & "=" & nm.RefersTo & ","\nNext nm\n'
                    'Report = Report & "|"\nEnd If\nNext bk\n')
            code = 'Function Report() As String\nDim src As Object, dst As Object, bk As Object, sh As Object, nm As Object\n' + body + 'End Function'
            result = excel.run_vba(code, 'Report', timeout=120)
            records.append(dict(name=name, body=body, ok=result.ok, reported=str(result.value), message=result.message))
            print(name, result.ok, result.value, result.message, flush=True)
    (Path(__file__).resolve().parents[1] / 'tests/fixtures/sheet_move.json').write_text(json.dumps(records, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
