"""Measure adding and deleting containing Forms group boxes."""
import json
from pathlib import Path
from zipfile import ZipFile

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent


def main() -> None:
    records = []
    folder = ROOT / ".cache/project-review/radio-regrouping"
    folder.mkdir(parents=True, exist_ok=True)
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for operation in ("add", "delete", "add_suffix", "delete_suffix", "add_all", "delete_all"):
            box_args = "190, 0, 150, 100" if "suffix" in operation else ("0, 0, 350, 100" if "all" in operation else "0, 0, 150, 100")
            make_box = f'Set box = ActiveSheet.Shapes.AddFormControl(4, {box_args})\n'
            for chosen in (0, 1, 2):
                for outside in (False, True):
                    body = 'Range("H1:H2").ClearContents\n'
                    if operation.startswith("delete"):
                        body += make_box
                    body += ('Set a = ActiveSheet.Shapes.AddFormControl(7, 10, 10, 90, 20)\n'
                             'Set b = ActiveSheet.Shapes.AddFormControl(7, 10, 40, 90, 20)\n'
                             'Set c = ActiveSheet.Shapes.AddFormControl(7, 200, 10, 90, 20)\n'
                             'a.Name = "First"\nb.Name = "Second"\nc.Name = "Third"\n'
                             'a.ControlFormat.LinkedCell = "$H$1"\n')
                    if operation.startswith("delete"):
                        body += 'c.ControlFormat.LinkedCell = "$H$2"\n'
                    if chosen:
                        body += f'{"a" if chosen == 1 else "b"}.ControlFormat.Value = 1\n'
                    if outside:
                        body += 'c.ControlFormat.Value = 1\n'
                    body += make_box if operation.startswith("add") else 'box.Delete\n'
                    report = ('CStr(a.ControlFormat.Value) & ":" & a.ControlFormat.LinkedCell & ";" & _\n'
                              'CStr(b.ControlFormat.Value) & ":" & b.ControlFormat.LinkedCell & ";" & _\n'
                              'CStr(c.ControlFormat.Value) & ":" & c.ControlFormat.LinkedCell & ";" & _\n'
                              'CStr(Range("H1").Value) & ":" & CStr(Range("H2").Value)')
                    body += 'Report = ' + report + '\n'
                    body += 'b.ControlFormat.Value = 1\nReport = Report & "|" & ' + report + '\n'
                    target = folder / "state.xlsm"
                    code = ('Public Function Report() As String\nDim a As Object, b As Object, c As Object, box As Object\n'
                            'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                            + body + 'Application.DisplayAlerts = False\n'
                            + f'ActiveWorkbook.SaveAs "{target}", 52\nEnd Function\n')
                    result = excel.run_vba(code, "Report", timeout=120.0)
                    assert result.ok, result.message
                    with ZipFile(target) as package:
                        parts = {part: package.read(part).decode() for part in package.namelist()
                                 if part.startswith("xl/ctrlProps/") or part.endswith(".vml")}
                    records.append({"operation": operation, "chosen": chosen, "outside": outside,
                                    "body": body, "reported": str(result.value), "parts": parts})
                    print(operation, chosen, outside, result.value)
    (ROOT / "tests/fixtures/shapes/radio_regrouping.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
