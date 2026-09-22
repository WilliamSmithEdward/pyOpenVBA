"""Measure multi/extended list link and source changes in Excel."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "link_empty": 'sh.ControlFormat.LinkedCell = "$H$2"',
    "link_numeric": 'Range("H2").Value = 2\nsh.ControlFormat.LinkedCell = "$H$2"',
    "link_text": 'Range("H2").Value = "text"\nsh.ControlFormat.LinkedCell = "$H$2"',
    "unlink": 'sh.ControlFormat.LinkedCell = ""',
    "cell_write": 'Range("H1").Value = 2',
    "source": 'sh.ControlFormat.ListFillRange = "$J$1:$J$2"',
    "source_short": 'sh.ControlFormat.ListFillRange = "$J$1"',
    "source_empty": 'sh.ControlFormat.ListFillRange = ""',
    "same_mode": 'sh.ControlFormat.MultiSelect = MODE',
    "other_mode": 'sh.ControlFormat.MultiSelect = OTHER',
    "single_mode": 'sh.ControlFormat.MultiSelect = 1',
    "source_expand": 'sh.ControlFormat.ListFillRange = "$J$1"\nsh.ControlFormat.ListFillRange = "$J$1:$J$3"',
    "source_clear": 'sh.ControlFormat.ListFillRange = "$J$1:$J$2"\nsh.ControlFormat.ListFillRange = ""',
    "scalar_rebind": 'sh.ControlFormat.MultiSelect = 1\nsh.ControlFormat.Value = 2\nsh.ControlFormat.MultiSelect = MODE\nRange("H2").Value = 9\nsh.ControlFormat.LinkedCell = "$H$2"',
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for mode in (2, 3):
            for name, statement in PROBES.items():
                statement = statement.replace("MODE", str(mode)).replace("OTHER", str(5 - mode))
                body = ('Range("H1:H2").ClearContents\nRange("J1").Value = "x"\nRange("J2").Value = "y"\n'
                        'Set sh = ActiveSheet.Shapes.AddFormControl(6, 0, 0, 90, 60)\n'
                        'sh.ControlFormat.List = Array("a", "b", "c")\n'
                        'sh.ControlFormat.LinkedCell = "$H$1"\n'
                        f'sh.ControlFormat.MultiSelect = {mode}\n'
                        'sh.DrawingObject.Selected(1) = True\nsh.DrawingObject.Selected(3) = True\n'
                        'On Error Resume Next\n' + statement + '\nn = Err.Number\nOn Error GoTo 0\n'
                        'Report = CStr(n) & "|" & sh.ControlFormat.LinkedCell & "|" & sh.ControlFormat.ListFillRange & "|" & _\n'
                        'CStr(Range("H1").Value) & "|" & CStr(Range("H2").Value) & "|" & CStr(sh.ControlFormat.MultiSelect) & "|"\n'
                        'For i = 1 To sh.ControlFormat.ListCount\n'
                        'Report = Report & sh.ControlFormat.List(i) & ":" & CStr(sh.DrawingObject.Selected(i)) & ";"\nNext i\n')
                code = ('Public Function Report() As String\nDim sh As Object, n As Long, i As Long\n'
                        'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                        + body + 'End Function\n')
                result = excel.run_vba(code, "Report", timeout=120.0)
                assert result.ok, f"{mode} {name}: {result.message}"
                records.append({"mode": mode, "name": name, "body": body, "reported": str(result.value)})
                print(mode, name, result.value)
    (ROOT / "tests/fixtures/shapes/list_rebinding.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
