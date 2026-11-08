"""Application.LanguageSettings and Application.VBE as a macro reads them, typed and late bound, in live Excel.

Each read runs in its own macro with two new workbooks open beside the harness's book, the second active, and answers the
read's TypeName and value, or the error it raised. Nothing is set.

Predictions, written before the run: LanguageSettings answers a LanguageSettings whose LanguageID is 1033 for the install, the
user interface, help and the previous interface, and 5 for anything else; LanguagePreferredForEditing is True for 1033 alone. VBE
answers the same VBE each read, Version 7.01; VBProjects holds each open workbook's project in the order the workbooks opened,
and ActiveVBProject is the project of the workbook whose macro is running. Nothing is selected in the editor, so
SelectedVBComponent and ActiveCodePane are Nothing, and the main window is hidden.

It only reads, so it takes a slot, not live.lock, and touches no member that could send mail or acts on the machine.

    python scripts/measure_application_language_vbe.py

writes tests/fixtures/application_language_vbe.json, which tests/test_excel_application_language_vbe.py replays.
"""
from __future__ import annotations

import json
import os
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/application_language_vbe.json"
READS = [
    "TypeName(a.LanguageSettings)", "a.LanguageSettings Is a.LanguageSettings", "a.LanguageSettings.Parent Is a",
    "a.LanguageSettings.Application Is a", "a.LanguageSettings.Creator",
    *(f"a.LanguageSettings.LanguageID({at})" for at in range(-1, 8)), 'a.LanguageSettings.LanguageID("2")',
    "a.LanguageSettings.LanguageID()", *(f"a.LanguageSettings.LanguagePreferredForEditing({at})" for at in (1033, 1031, 1041, 0, -1, 2057)),
    "a.LanguageSettings.LanguagePreferredForEditing()", "a.LanguageSettings.LanguageID", "a.LanguageSettings.Count",
    "TypeName(a.VBE)", "a.VBE Is a.VBE", "a.VBE.Version", "TypeName(a.VBE.VBProjects)", "a.VBE.VBProjects.Count",
    "a.VBE.VBProjects(1) Is ThisWorkbook.VBProject", "a.VBE.VBProjects(a.VBE.VBProjects.Count) Is ActiveWorkbook.VBProject",
    "a.VBE.VBProjects(a.VBE.VBProjects.Count - 1) Is Workbooks(Workbooks.Count - 1).VBProject", "a.VBE.VBProjects(0)",
    "a.VBE.VBProjects(99)", 'a.VBE.VBProjects("VBAProject") Is ThisWorkbook.VBProject', 'a.VBE.VBProjects("Nope")',
    "a.VBE.VBProjects.Item(1) Is ThisWorkbook.VBProject", "a.VBE.VBProjects.Parent Is a.VBE", "a.VBE.VBProjects.VBE Is a.VBE",
    "TypeName(a.VBE.ActiveVBProject)", "a.VBE.ActiveVBProject Is ThisWorkbook.VBProject", "a.VBE.ActiveVBProject Is ActiveWorkbook.VBProject",
    "a.VBE.ActiveVBProject.Name", "ThisWorkbook.VBProject.VBE Is a.VBE", "ThisWorkbook.VBProject.Collection Is a.VBE.VBProjects",
    "TypeName(ThisWorkbook.VBProject.Collection)", "ActiveWorkbook.VBProject.Name", "a.VBE.SelectedVBComponent Is Nothing",
    "a.VBE.ActiveCodePane Is Nothing", "TypeName(a.VBE.MainWindow)", "a.VBE.MainWindow.Visible", "a.VBE.Windows.Count",
    "a.VBE.CodePanes.Count", "TypeName(a.VBE.Addins)", "a.VBE.Addins.Count", "a.VBE.Parent", "a.VBE.Name", "a.VBE.Count",
    "TypeName(a.VBE.Events)", "TypeName(a.VBE.CommandBars)", "a.VBE.ActiveWindow Is Nothing", "a.VBE.VBProjects(1).Name",
    "a.VBE.VBProjects(2).Name", "a.VBE.VBProjects(1).VBComponents.Count > 0", "Walk(a)",
]
SOURCE = """Function Walk(a As Object) As String
    Dim p As Object, s As String
    For Each p In a.VBE.VBProjects
        s = s & p.Name & "/" & (p Is ThisWorkbook.VBProject) & "/" & (p Is ActiveWorkbook.VBProject) & ";"
    Next
    Walk = s
End Function

Function Go() As String
    {dim}
    Dim v As Variant
    Set a = Application
    On Error Resume Next
    v = {read}
    If Err.Number <> 0 Then
        Go = "E" & Err.Number
    Else
        Go = TypeName(v) & ":" & CStr(v)
    End If
End Function
"""
BINDINGS = {"typed": "Dim a As Application", "late": "Dim a As Object"}
#: Reads that do not compile bound early, which stops the harness: asked late bound only.
LATE_ONLY = {"a.LanguageSettings.LanguageID()", "a.LanguageSettings.LanguagePreferredForEditing()", "a.LanguageSettings.LanguageID",
             "a.LanguageSettings.Count", "a.VBE.Parent", "a.VBE.Name", "a.VBE.Count"}


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    common = Path(subprocess.run(["git", "rev-parse", "--git-common-dir"], cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip())
    base = (common if common.is_absolute() else ROOT / common) / "excel-coordination"
    slot = None
    for _ in range(900):
        if not (base / "live.lock").exists():
            for number in (1, 2, 3):
                try:
                    os.mkdir(base / f"live-slot-{number}")
                    slot = base / f"live-slot-{number}"
                    break
                except FileExistsError:
                    continue
        if slot is not None:
            break
        time.sleep(1)
    else:
        raise TimeoutError("no slot came free")
    rows: dict[str, dict[str, str]] = {}
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=900.0, exclusive=False)) as excel:
            excel.new_document()
            opened = excel.run_vba('Function Go() As String\n    Workbooks.Add\n    Workbooks.Add\n    Go = Workbooks.Count & ":" & ActiveWorkbook.Name\nEnd Function\n',
                                   "Go", (), timeout=60.0, line_numbers=False)
            assert opened.ok, opened
            rows["opened"] = {"late": str(opened.value)}
            for read in READS:
                rows[read] = {}
                for binding, dim in BINDINGS.items():
                    if binding == "typed" and read in LATE_ONLY:
                        continue
                    answer = excel.run_vba(SOURCE.format(dim=dim, read=read), "Go", (), timeout=60.0, line_numbers=False)
                    rows[read][binding] = str(answer.value) if answer.ok else f"[{answer.outcome}]"
                print(read, rows[read], flush=True)
    finally:
        os.rmdir(slot)
    OUT.write_text(json.dumps({"rows": rows, "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
