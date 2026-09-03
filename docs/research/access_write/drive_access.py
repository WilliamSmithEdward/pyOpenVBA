"""Drive Access itself through the harness, so a module operation runs
inside Access's own VBA rather than over a COM boundary that refuses
half the DoCmd verbs.

The same source is injected for every run and only the procedure called
differs, so two databases driven this way differ by the operation alone.
"""

from __future__ import annotations

import sys
from pathlib import Path

from pyvbaharness import AccessSession

SOURCE = """
Public Function NoOp() As Variant
    NoOp = "ok"
End Function

Public Function DoRename(ByVal oldName As String, ByVal newName As String) As Variant
    DoCmd.Rename newName, acModule, oldName
    DoRename = "ok"
End Function

Public Function DoAdd(ByVal name As String) As Variant
    Dim c As Object
    Set c = Application.VBE.ActiveVBProject.VBComponents.Add(1)
    c.Name = name
    c.CodeModule.AddFromString "Public Function " & name & "Go() As Variant" & vbCrLf & _
        "    " & name & "Go = 42" & vbCrLf & "End Function"
    DoCmd.Save acModule, name
    DoAdd = "ok"
End Function

Public Function DoDelete(ByVal name As String) As Variant
    DoCmd.DeleteObject acModule, name
    DoDelete = "ok"
End Function

Public Function DoAddBody(ByVal name As String, ByVal body As String) As Variant
    Dim c As Object
    Set c = Application.VBE.ActiveVBProject.VBComponents.Add(1)
    c.Name = name
    c.CodeModule.AddFromString Replace(body, "|", vbCrLf)
    DoCmd.Save acModule, name
    DoAddBody = "ok"
End Function

Public Function Compiles() As Variant
    On Error GoTo bad
    Application.VBE.ActiveVBProject.VBComponents(1).CodeModule.CountOfLines
    DoCmd.RunCommand 126
    Compiles = "compiled"
    Exit Function
bad:
    Compiles = "error " & Err.Number & ": " & Err.Description
End Function

Public Function EvalIt(ByVal expression As String) As Variant
    EvalIt = Eval(expression)
End Function

Public Function AddAndRun(ByVal moduleName As String) As Variant
    Dim c As Object
    Set c = Application.VBE.ActiveVBProject.VBComponents(moduleName)
    c.CodeModule.AddFromString "Public Function Probe9() As Variant" & vbCrLf & _
        "    Probe9 = 999" & vbCrLf & "End Function"
    AddAndRun = Application.Run("Probe9")
End Function

Public Function AccessModules() As Variant
    Dim i As Long, s As String
    For i = 0 To CurrentProject.AllModules.Count - 1
        s = s & CurrentProject.AllModules(i).Name & ";"
    Next i
    AccessModules = s
End Function

Public Function CallProc(ByVal name As String) As Variant
    CallProc = Application.Run(name)
End Function

Public Function ModuleNames() As Variant
    Dim i As Long, s As String
    For i = 1 To Application.VBE.ActiveVBProject.VBComponents.Count
        s = s & Application.VBE.ActiveVBProject.VBComponents(i).Name & ";"
    Next i
    ModuleNames = s
End Function
"""


def run(path: Path, proc: str, args: tuple = ()) -> object:
    with AccessSession() as access:
        access.open_document(path, read_only=False)
        result = access.run_vba(SOURCE, proc=proc, args=args, timeout=120)
        if result.outcome != "passed":
            raise SystemExit(f"{proc} {args}: {result.outcome} {getattr(result, 'error', None)}")
        return result.value


if __name__ == "__main__":
    target = Path(sys.argv[1])
    print(run(target, sys.argv[2], tuple(sys.argv[3:])))
