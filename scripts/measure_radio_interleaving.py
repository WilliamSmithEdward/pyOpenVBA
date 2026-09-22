"""Measure interleaved radio creation and the resulting package order."""
import json
from pathlib import Path
from zipfile import ZipFile

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent


def main() -> None:
    folder = ROOT / ".cache/project-review/radio-interleaving"
    folder.mkdir(parents=True, exist_ok=True)
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for boxes in (1, 2):
            for order in ("ABAB", "ABBA", "BAAB", "ABCABC"):
                for operation in ("select", "delete_first_box", "delete_second_box", "add_first_box",
                                  "add_first_box_last_a", "add_first_box_last_b"):
                    if boxes == 1 and operation == "delete_second_box":
                        continue
                    body = 'Range("H1:H2").ClearContents\n'
                    if not operation.startswith("add_first_box"):
                        body += 'Set box1 = ActiveSheet.Shapes.AddFormControl(4, 0, 0, 150, 200)\n'
                    if boxes == 2:
                        body += 'Set box2 = ActiveSheet.Shapes.AddFormControl(4, 200, 0, 150, 200)\n'
                    for i, zone in enumerate(order, 1):
                        x = {"A": 10, "B": 210, "C": 410}[zone]
                        body += (f'Set sh = ActiveSheet.Shapes.AddFormControl(7, {x}, {i * 25}, 90, 20)\n'
                                 f'sh.Name = "Radio{i}"\n')
                    a = order.index("A") + 1
                    b = order.index("B") + 1
                    body += (f'ActiveSheet.Shapes("Radio{a}").ControlFormat.LinkedCell = "$H$1"\n'
                             f'ActiveSheet.Shapes("Radio{b}").ControlFormat.LinkedCell = "$H$2"\n'
                             f'ActiveSheet.Shapes("Radio{a}").ControlFormat.Value = 1\n'
                             f'ActiveSheet.Shapes("Radio{b}").ControlFormat.Value = 1\n')
                    if operation.endswith("last_a") or operation.endswith("last_b"):
                        index = order.rindex("A" if operation.endswith("last_a") else "B") + 1
                        body += f'ActiveSheet.Shapes("Radio{index}").ControlFormat.Value = 1\n'
                    if operation.startswith("add_first_box"):
                        body += 'Set box1 = ActiveSheet.Shapes.AddFormControl(4, 0, 0, 150, 200)\n'
                    elif operation != "select":
                        body += 'box1.Delete\n' if operation == "delete_first_box" else 'box2.Delete\n'
                    body += ('Report = CStr(Range("H1").Value) & ":" & CStr(Range("H2").Value) & "|"\n'
                             'For Each sh In ActiveSheet.Shapes\nIf Left(sh.Name, 5) = "Radio" Then\n'
                             'Report = Report & sh.Name & ":" & CStr(sh.ControlFormat.Value) & ":" & sh.ControlFormat.LinkedCell & ";"\n'
                             'End If\nNext sh\n')
                    target = folder / "state.xlsm"
                    code = ('Public Function Report() As String\nDim sh As Object, box1 As Object, box2 As Object\n'
                            'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                            + body + 'Application.DisplayAlerts = False\n'
                            + f'ActiveWorkbook.SaveAs "{target}", 52\nEnd Function\n')
                    result = excel.run_vba(code, "Report", timeout=120.0)
                    assert result.ok, result.message
                    with ZipFile(target) as package:
                        parts = {part: package.read(part).decode() for part in package.namelist()
                                 if part.startswith("xl/ctrlProps/") or part.endswith(".vml")
                                 or part in ("xl/drawings/drawing1.xml", "xl/worksheets/sheet1.xml")}
                    records.append({"boxes": boxes, "order": order, "operation": operation,
                                    "body": body, "reported": str(result.value), "parts": parts})
                    print(boxes, order, operation, result.value)
    (ROOT / "tests/fixtures/shapes/radio_interleaving.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
