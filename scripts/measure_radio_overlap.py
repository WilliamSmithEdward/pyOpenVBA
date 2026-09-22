"""Measure overlapping group-box ownership and regrouping in Excel."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for layout in ("overlap", "overlap_small", "nested", "identical"):
            first = "0, 0, 240, 180"
            second = {"overlap": "120, 0, 240, 180", "overlap_small": "120, 0, 180, 180", "nested": "60, 0, 120, 180",
                      "identical": first}[layout]
            for reverse in (False, True):
                for operation in ("select", "delete_first", "delete_second", "late_second"):
                    create_first = f'Set box1 = ActiveSheet.Shapes.AddFormControl(4, {first})\n'
                    create_second = f'Set box2 = ActiveSheet.Shapes.AddFormControl(4, {second})\n'
                    body = 'Range("H1:H3").ClearContents\n'
                    if operation == "late_second":
                        body += create_first
                    else:
                        body += create_second + create_first if reverse else create_first + create_second
                    for i, x in enumerate((10, 130, 250), 1):
                        body += (f'Set sh = ActiveSheet.Shapes.AddFormControl(7, {x}, 30, 40, 20)\n'
                                 f'sh.Name = "Radio{i}"\n')
                    for i in (1, 2, 3):
                        body += f'ActiveSheet.Shapes("Radio{i}").ControlFormat.LinkedCell = "$H${i}"\n'
                    for i in (1, 2, 3):
                        body += f'ActiveSheet.Shapes("Radio{i}").ControlFormat.Value = 1\n'
                    if operation == "delete_first":
                        body += 'box1.Delete\n'
                    elif operation == "delete_second":
                        body += 'box2.Delete\n'
                    elif operation == "late_second":
                        body += create_second
                    body += ('Report = CStr(Range("H1").Value) & ":" & CStr(Range("H2").Value) & ":" & CStr(Range("H3").Value) & "|"\n'
                             'For Each sh In ActiveSheet.Shapes\nIf Left(sh.Name, 5) = "Radio" Then\n'
                             'Report = Report & sh.Name & ":" & CStr(sh.ControlFormat.Value) & ":" & sh.ControlFormat.LinkedCell & ";"\n'
                             'End If\nNext sh\n')
                    code = ('Public Function Report() As String\nDim sh As Object, box1 As Object, box2 As Object\n'
                            'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                            + body + 'End Function\n')
                    result = excel.run_vba(code, "Report", timeout=120.0)
                    assert result.ok, result.message
                    records.append({"layout": layout, "reverse": reverse, "operation": operation,
                                    "body": body, "reported": str(result.value)})
                    print(layout, reverse, operation, result.value)
    (ROOT / "tests/fixtures/shapes/radio_overlap.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
