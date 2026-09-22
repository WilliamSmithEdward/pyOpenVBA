"""Measure scoped Excel name CRUD."""
import json
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

PROBES = {
    "create": 'Set nm = ThisWorkbook.Names.Add(Name:="TestName", RefersTo:="=Sheet1!$A$1", Visible:=False)\nanswer = nm.Name & ":" & nm.RefersTo & ":" & CStr(nm.Visible)',
    "local": 'Set nm = Worksheets("Sheet1").Names.Add(Name:="TestName", RefersTo:="=Sheet1!$B$1")\nanswer = nm.Name & ":" & CStr(Worksheets("Sheet1").Names.Count)',
    "duplicate": 'ThisWorkbook.Names.Add "TestName", "=Sheet1!$B$1"\nanswer = CStr(ThisWorkbook.Names.Count) & ":" & ThisWorkbook.Names("TestName").RefersTo',
    "rename": 'ThisWorkbook.Names("TestName").Name = "Renamed"\nanswer = Range("C1").Formula',
    "delete": 'ThisWorkbook.Names("TestName").Delete\nanswer = Range("C1").Formula & ":" & CStr(Range("C1").Value)',
    "retarget": 'ThisWorkbook.Names("TestName").RefersTo = "=Sheet1!$B$1"\nanswer = CStr(Range("C1").Value)',
    "value_alias": 'ThisWorkbook.Names("TestName").Value = "=Sheet1!$B$1"\nanswer = ThisWorkbook.Names("TestName").RefersTo',
    "comment": 'ThisWorkbook.Names("TestName").Comment = "note"\nanswer = ThisWorkbook.Names("TestName").Comment',
    "collision": 'ThisWorkbook.Names.Add "OtherName", "=Sheet1!$B$1"\nThisWorkbook.Names("TestName").Name = "OtherName"\nanswer = ThisWorkbook.Names(1).Name & ":" & ThisWorkbook.Names(1).RefersTo & "|" & ThisWorkbook.Names(2).Name & ":" & ThisWorkbook.Names(2).RefersTo',
    "bad_cell": 'ThisWorkbook.Names.Add "A1", "=Sheet1!$A$1"',
    "bad_space": 'ThisWorkbook.Names.Add "bad name", "=Sheet1!$A$1"',
    "case": 'ThisWorkbook.Names.Add "testname", "=Sheet1!$B$1"\nanswer = ThisWorkbook.Names("TestName").Name',
    "scope_lookup": 'Worksheets("Sheet1").Names.Add "TestName", "=Sheet1!$B$1"\nanswer = CStr(Range("TestName").Value) & ":" & CStr(Range("C1").Value) & ":" & ThisWorkbook.Names("TestName").Name & ":" & Worksheets("Sheet1").Names("TestName").Name',
    "range_set_name": 'Range("B1").Name = "NewName"\nanswer = ThisWorkbook.Names("NewName").RefersTo',
    "r1c1": 'Set nm = ThisWorkbook.Names.Add(Name:="OtherName", RefersToR1C1:="=Sheet1!R1C2")\nanswer = nm.RefersTo & ":" & nm.RefersToR1C1',
    "range_get_name": 'answer = Range("A1").Name.Name',
    "missing": 'Set nm = ThisWorkbook.Names("Missing")',
    "missing_range_name": 'answer = Range("B1").Name.Name',
    "index_order": 'ThisWorkbook.Names.Add "Zed", "=1"\nThisWorkbook.Names.Add "Alpha", "=2"\nanswer = ThisWorkbook.Names(1).Name & ":" & ThisWorkbook.Names(2).Name & ":" & ThisWorkbook.Names(3).Name',
}

def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        for name, statement in PROBES.items():
            excel.new_document()
            body = ('Range("A1").Value = 1\nRange("B1").Value = 2\n'
                    'ThisWorkbook.Names.Add "TestName", "=Sheet1!$A$1"\nRange("C1").Formula = "=TestName"\n'
                    'answer = CStr(Range("C1").Value)\nanswer = ""\nOn Error Resume Next\n' + statement + '\nn = Err.Number\nOn Error GoTo 0\n'
                    'Report = CStr(n) & "|" & answer & "|" & CStr(ThisWorkbook.Names.Count)\n')
            result = excel.run_vba('Function Report() As String\nDim nm As Object, n As Long, answer As String\n' + body + 'End Function', "Report", timeout=120.0)
            assert result.ok, result.message
            records.append(dict(name=name, body=body, reported=str(result.value)))
            print(name, result.value, flush=True)
    (Path(__file__).resolve().parents[1] / "tests/fixtures/names_crud.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")

if __name__ == "__main__":
    main()
