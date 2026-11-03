"""Objects and values Application hands a macro that has no screen: AutoRecover, MultiThreadedCalculation, AddIns,
RecentFiles and the rest. Read only: nothing here changes the measuring user's profile, and a value that names the
machine or its user is reported by its shape, never its text.

    python scripts/measure_application_objects.py

writes tests/fixtures/application_objects.json, which tests/test_excel_application_objects.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/application_objects.json"

#: Members read straight from Application: the type each answers and, for what names no one, the value.
PLAIN = ["AutoRecover", "MultiThreadedCalculation", "AddIns", "AddIns2", "COMAddIns", "RecentFiles", "Dialogs",
         "CommandBars", "AutoCorrect", "SpellingOptions", "DefaultWebOptions", "LanguageSettings", "Assistance",
         "FileValidation", "FileValidationPivot", "FileConverters", "FileExportConverters", "SmartArtLayouts",
         "SmartArtQuickStyles", "SmartArtColors", "Speech", "UsedObjects", "WindowsForPens", "VBE", "Watches",
         "ActiveEncryptionSession", "ClusterConnector", "DefaultPivotTableLayoutOptions", "RegisteredFunctions",
         "ClipboardFormats", "DDEAppReturnCode", "MailSession", "MailSystem", "AlwaysUseClearType",
         "DisplayXMLSourcePane", "FormulaBarHeight", "SensitivityLabelPolicy", "QuickAnalysis",
         "SharePointVersion", "Build", "DisplayAlerts"]
#: Text that names this machine or its user: only whether it is there.
SHAPES = ["Path", "StartupPath", "AltStartupPath", "TemplatesPath", "NetworkTemplatesPath", "LibraryPath",
          "UserLibraryPath", "DefaultFilePath", "OrganizationName", "ProductCode", "UserName", "ActivePrinter",
          "Hwnd", "Hinstance", "HinstancePtr"]
#: Further reads through an object: object expression, members.
THROUGH = {
    "AutoRecover": ["Enabled", "Time", "Application", "Creator", "Parent"],
    "MultiThreadedCalculation": ["Enabled", "Mode", "Application", "Creator", "Parent"],
    "AddIns": ["Count", "Application", "Creator", "Parent"],
    "AddIns2": ["Count"],
    "RecentFiles": ["Maximum", "Application", "Creator", "Parent"],
    "Dialogs": ["Count", "Application", "Creator", "Parent"],
    "SpellingOptions": ["ArabicModes", "IgnoreCaps", "IgnoreMixedDigits", "IgnoreFileNames", "SuggestMainOnly",
                        "UserDict", "DictLang"],
    "AutoCorrect": ["CapitalizeNamesOfDays", "CorrectCapsLock", "CorrectSentenceCap", "ReplaceText", "TwoInitialCapitals",
                    "AutoExpandListRange", "AutoFillFormulasInLists", "DisplayAutoCorrectOptions", "Application",
                    "Creator", "Parent"],
    "DefaultWebOptions": ["AllowPNG", "DownloadComponents", "Encoding", "LocationOfComponents", "OrganizeInFolder",
                          "PixelsPerInch", "RelyOnCSS", "RelyOnVML", "SaveHiddenData", "SaveNewWebPagesAsWebArchives",
                          "ScreenSize", "TargetBrowser", "UpdateLinksOnSave", "UseLongFileNames", "AlwaysSaveInDefaultEncoding",
                          "CheckIfOfficeIsHTMLEditor", "LoadPictures"],
    "DefaultPivotTableLayoutOptions": ["Application", "Creator", "Parent"],
    "FileConverters": [],
}

PROBE = '''Private Function Fetch(owner As Object, ByVal name As String) As String
    On Error GoTo Failed
    Dim v As Variant
    If IsObject(CallByName(owner, name, VbGet)) Then
        Fetch = "object:" & TypeName(CallByName(owner, name, VbGet))
    Else
        v = CallByName(owner, name, VbGet)
        If IsArray(v) Then
            Fetch = "array:" & TypeName(v)
        Else
            Fetch = TypeName(v) & ":" & CStr(v)
        End If
    End If
    Exit Function
Failed:
    Fetch = "E" & Err.Number
End Function

Private Function Shape(owner As Object, ByVal name As String) As String
    On Error GoTo Failed
    Dim v As Variant
    v = CallByName(owner, name, VbGet)
    Shape = TypeName(v) & ":" & IIf(Len(CStr(v)) > 0, "something", "empty")
    Exit Function
Failed:
    Shape = "E" & Err.Number
End Function

Private Function Alerts(ByVal given As Variant) As String
    Dim was As Boolean
    was = Application.DisplayAlerts
    On Error Resume Next
    Application.DisplayAlerts = given
    Alerts = Application.DisplayAlerts & "/E" & Err.Number
    On Error GoTo 0
    Application.DisplayAlerts = was
End Function

Public Function Direct() As String
    Dim s As String
    s = TypeName(Application.Build) & "|" & (Application.Build > 0) & "|" & TypeName(Application.ThisWorkbook)
    s = s & "|" & (Application.ThisWorkbook Is ThisWorkbook)
    Application.DisplayAlerts = False
    s = s & "|" & Alerts(True) & "|" & Alerts(False) & "|" & Alerts(0) & "|" & Alerts(1) & "|" & Alerts("True")
    Application.DisplayAlerts = True
    s = s & "|" & Alerts("x") & "|" & Alerts(Null) & "|" & Alerts(Empty)
    Application.DisplayAlerts = False
    Direct = s
End Function

Public Function Probe() As String
    Dim s As String, o As Object
{lines}
    Probe = Mid$(s, 4)
End Function
'''


def probe() -> str:
    lines = [f'    s = s & "@@@{name}=" & Fetch(Application, "{name}")' for name in PLAIN]
    lines += [f'    s = s & "@@@{name}=" & Shape(Application, "{name}")' for name in SHAPES]
    for owner, members in THROUGH.items():
        lines += ['    Set o = Nothing', '    On Error Resume Next', f'    Set o = CallByName(Application, "{owner}", VbGet)',
                  '    On Error GoTo 0']
        lines += [f'    If Not o Is Nothing Then s = s & "@@@{owner}.{name}=" & Fetch(o, "{name}")' for name in members]
    return PROBE.format(lines="\n".join(lines))


def main() -> None:
    source = probe()
    with ExcelSession(HarnessConfig(lock_wait_s=60.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, "Probe", timeout=120.0)
        assert result.ok, result
        direct = excel.run_vba(source, "Direct", timeout=60.0)
        assert direct.ok, direct
    print("direct", direct.value, flush=True)
    rows = dict(item.split("=", 1) for item in str(result.value).split("@@@"))
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": source, "direct": str(direct.value),
                               "rows": rows},
                              indent=1) + "\n",
                   encoding="utf-8")
    for name, value in rows.items():
        print(name, value, flush=True)


if __name__ == "__main__":
    main()
