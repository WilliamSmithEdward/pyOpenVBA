"""Measure option-button deletion and group-link transfer in Excel."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for boxed in (False, True):
            for selected in (1, 2, 3):
                for deleted in (1, 2, 3):
                    body = ('Range("H1").ClearContents\n'
                            + ('ActiveSheet.Shapes.AddFormControl 4, 0, 0, 150, 100\n' if boxed else '')
                            + 'For i = 1 To 3\nSet sh = ActiveSheet.Shapes.AddFormControl(7, 10, i * 25, 90, 20)\n'
                            'sh.Name = "Radio" & CStr(i)\nNext i\n'
                            'ActiveSheet.Shapes("Radio1").ControlFormat.LinkedCell = "$H$1"\n'
                            f'ActiveSheet.Shapes("Radio{selected}").ControlFormat.Value = 1\n'
                            f'ActiveSheet.Shapes("Radio{deleted}").Delete\n'
                            'Report = CStr(Range("H1").Value) & "|"\n'
                            'For Each sh In ActiveSheet.Shapes\nIf Left(sh.Name, 5) = "Radio" Then\n'
                            'Report = Report & sh.Name & ":" & CStr(sh.ControlFormat.Value) & ":" & sh.ControlFormat.LinkedCell & ";"\n'
                            'End If\nNext sh\n')
                    code = ('Public Function Report() As String\nDim sh As Object, i As Long\n'
                            'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                            + body + 'End Function\n')
                    result = excel.run_vba(code, "Report", timeout=120.0)
                    assert result.ok, result.message
                    records.append({"boxed": boxed, "selected": selected, "deleted": deleted,
                                    "body": body, "reported": str(result.value)})
                    print(boxed, selected, deleted, result.value)
    (ROOT / "tests/fixtures/shapes/radio_deletion.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
