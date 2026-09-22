"""Record live Excel control values and the corresponding saved XML."""
from pathlib import Path
import json
from zipfile import ZipFile

from pyvbaharness import ExcelSession

ROOT = Path(__file__).resolve().parent.parent


def main() -> None:
    folder = ROOT / ".cache" / "project-review" / "control-values"
    folder.mkdir(parents=True, exist_ok=True)
    reports = []
    with ExcelSession() as excel:
        excel.new_document()
        for index, value in enumerate((-4146, 1, 2)):
            target = folder / f"state-{index}.xlsm"
            setup = ('Set sh = ActiveSheet.Shapes.AddFormControl(1, 10, 10, 90, 20)\n'
                     'sh.Name = "Check"\n') if index == 0 else 'Set sh = ActiveSheet.Shapes("Check")\n'
            code = (
                'Public Function Report() As String\nDim sh As Object\n' + setup
                + f'sh.ControlFormat.Value = {value}\n'
                + 'Report = CStr(sh.ControlFormat.Value)\n'
                + f'ActiveWorkbook.SaveAs "{target}", 52\nEnd Function\n'
            )
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, f"{result.outcome}: {result.message}"
            with ZipFile(target) as package:
                parts = {name: package.read(name).decode() for name in package.namelist()
                         if name.startswith("xl/ctrlProps/") or name.endswith(".vml")}
            reports.append({"assigned": value, "reported": result.value, "parts": parts})
    (folder / "measured.json").write_text(json.dumps(reports, indent=2), encoding="utf-8")
    compact = [{"assigned": record["assigned"], "reported": record["reported"],
                "properties": record["parts"]["xl/ctrlProps/ctrlProp1.xml"]} for record in reports]
    (ROOT / "tests" / "fixtures" / "shapes" / "checkbox_values.json").write_text(
        json.dumps(compact, indent=2) + "\n", encoding="utf-8",
    )
    print(json.dumps(reports, indent=2))


if __name__ == "__main__":
    main()
