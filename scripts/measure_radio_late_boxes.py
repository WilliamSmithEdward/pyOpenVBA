"""Measure shared-link ownership when an overlapping box is added late."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for layout, dimensions in (("identical", "0, 0, 150, 100"), ("identical_all", "0, 0, 150, 100"), ("nested", "0, 0, 120, 35"),
                                   ("outer", "0, 0, 320, 120"), ("partial", "0, 30, 320, 100")):
            for links in ("none", "inside", "outside", "both"):
                for chosen in (0, 1, 2, 3):
                    body = ('Range("H1:H2").ClearContents\n'
                            'ActiveSheet.Shapes.AddFormControl 4, 0, 0, 150, 100\n'
                            'Set a = ActiveSheet.Shapes.AddFormControl(7, 10, 10, 90, 20)\n'
                            'Set b = ActiveSheet.Shapes.AddFormControl(7, 10, 40, 90, 20)\n'
                            + ('Set c = ActiveSheet.Shapes.AddFormControl(7, 10, 70, 90, 20)\n' if layout == "identical_all" else
                               'Set c = ActiveSheet.Shapes.AddFormControl(7, 200, 10, 90, 20)\n') +
                            'a.Name = "First"\nb.Name = "Second"\nc.Name = "Third"\n')
                    if links in ("inside", "both"):
                        body += 'a.ControlFormat.LinkedCell = "$H$1"\n'
                    if links in ("outside", "both"):
                        body += 'c.ControlFormat.LinkedCell = "$H$2"\n'
                    if chosen:
                        body += f'{"abc"[chosen - 1]}.ControlFormat.Value = 1\n'
                    body += f'ActiveSheet.Shapes.AddFormControl 4, {dimensions}\n'
                    report = ('CStr(a.ControlFormat.Value) & ":" & a.ControlFormat.LinkedCell & ";" & _\n'
                              'CStr(b.ControlFormat.Value) & ":" & b.ControlFormat.LinkedCell & ";" & _\n'
                              'CStr(c.ControlFormat.Value) & ":" & c.ControlFormat.LinkedCell & ";" & _\n'
                              'CStr(Range("H1").Value) & ":" & CStr(Range("H2").Value)')
                    body += 'Report = ' + report + '\n'
                    body += 'b.ControlFormat.Value = 1\nc.ControlFormat.Value = 1\nReport = Report & "|" & ' + report + '\n'
                    code = ('Public Function Report() As String\nDim a As Object, b As Object, c As Object\n'
                            'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                            + body + 'End Function\n')
                    result = excel.run_vba(code, "Report", timeout=120.0)
                    assert result.ok, result.message
                    records.append({"layout": layout, "links": links, "chosen": chosen,
                                    "body": body, "reported": str(result.value)})
                    print(layout, links, chosen, result.value)
    (ROOT / "tests/fixtures/shapes/radio_late_boxes.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
