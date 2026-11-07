"""Settable properties of Worksheet, Workbook and the Application settings macros set most, in live Excel: each value set
on a new workbook, typed and late bound, since a bad value's error differs between them, with its error and what the
property reads before and after. A case's answer is "before|Eset|after"; the first half of a row's cases is late bound. Application's settings are read first and put back, under live.lock: run with "application" for
those, and "books" for the rest; each writes its own part of the fixture.

    python scripts/measure_settable_audit.py books
    python scripts/measure_settable_audit.py application
    python scripts/measure_settable_audit.py sort

writes tests/fixtures/settable_audit.json, which scripts/compare_settable_audit.py and tests/test_excel_settable_audit.py
replay. AutoSaveOn is left out, as it can reach the cloud, and Colors, which takes an index. WorkIdentity is no longer set: no live
probe touches it, with rights, encryption and signing (tests/test_probes_leave_crypto_alone.py). The fixture keeps what an earlier run
recorded for it: do not remeasure it.

EnvelopeVisible, Password and WritePassword are no longer set: no live probe sets a user-level setting it does not put back, a mail envelope, the cloud save or a password
(tests/test_probes_leave_user_settings_alone.py). The fixture keeps what an earlier run recorded for them: do not remeasure them.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/settable_audit.json"

BOOLEAN = ["True", "False", "0", "1", "-1", "2", "1.5", '"True"', '"x"', '""', "Null", "Empty"]
NUMBER = ["0", "1", "2", "3", "-1", "1.5", "99999", "True", '"1"', '"x"', '""', "Null", "Empty"]
TEXT = ['"abc"', '""', "5", "1.5", "True", "Null", "Empty", '"' + "x" * 40 + '"']
VARIANT = sorted({*BOOLEAN, *TEXT})
WORKSHEET = {
    "_CodeName": TEXT, "AutoFilterMode": BOOLEAN, "DisplayPageBreaks": BOOLEAN, "DisplayRightToLeft": BOOLEAN,
    "EnableAutoFilter": BOOLEAN, "EnableCalculation": BOOLEAN, "EnableFormatConditionsCalculation": BOOLEAN,
    "EnableOutlining": BOOLEAN, "EnablePivotTable": BOOLEAN, "EnableSelection": NUMBER + ["-4142", "-4160"],
    "Name": TEXT + ['"a/b"', '"Sheet9"', '"Other"', '"History"'], "ScrollArea": TEXT + ['"A1:B2"', '"$A$1"', '"A1,C3"', '"Other!A1"'],
    "StandardWidth": NUMBER + ["255", "256", "8.43"], "TransitionExpEval": BOOLEAN, "TransitionFormEntry": BOOLEAN,
    "Visible": NUMBER + ["-1", "xlSheetVeryHidden", "xlSheetHidden"],
}
WORKBOOK = {
    "AccuracyVersion": NUMBER, "AutoUpdateFrequency": NUMBER, "AutoUpdateSaveChanges": BOOLEAN,
    "ChangeHistoryDuration": NUMBER, "ChartDataPointTrack": BOOLEAN, "CheckCompatibility": BOOLEAN,
    "CompatibilityVersion": NUMBER, "ConflictResolution": NUMBER, "Date1904": BOOLEAN,
    "DefaultPivotTableStyle": VARIANT, "DefaultSlicerStyle": VARIANT, "DefaultTableStyle": VARIANT,
    "DefaultTimelineStyle": VARIANT, "DisplayDrawingObjects": NUMBER + ["-4104", "-4142"], "DisplayInkComments": BOOLEAN,
    "DoNotPromptForConvert": BOOLEAN, "EnableAutoRecover": BOOLEAN,
    "ExternalCodeServiceTimeout": NUMBER, "Final": BOOLEAN, "ForceFullCalculation": BOOLEAN,
    "HighlightChangesOnScreen": BOOLEAN, "InactiveListBorderVisible": BOOLEAN, "IsAddin": BOOLEAN, "KeepChangeHistory": BOOLEAN,
    "ListChangesOnNewSheet": BOOLEAN, "PersonalViewListSettings": BOOLEAN,
    "PersonalViewPrintSettings": BOOLEAN, "PrecisionAsDisplayed": BOOLEAN, "ReadOnlyRecommended": BOOLEAN,
    "RemovePersonalInformation": BOOLEAN, "Saved": BOOLEAN, "SaveLinkValues": BOOLEAN, "ShowConflictHistory": BOOLEAN,
    "ShowPivotChartActiveFields": BOOLEAN, "ShowPivotTableFieldList": BOOLEAN, "TemplateRemoveExtData": BOOLEAN,
    "UpdateLinks": NUMBER, "UpdateRemoteReferences": BOOLEAN,
    "EncryptionProvider": TEXT,
}
APPLICATION = {
    "ScreenUpdating": BOOLEAN, "DisplayAlerts": BOOLEAN, "EnableEvents": BOOLEAN,
    "Calculation": NUMBER + ["-4105", "-4135", "2"], "Cursor": NUMBER + ["-4143", "xlWait", "xlIBeam", "xlNorthwestArrow"],
    "StatusBar": sorted({*VARIANT, "False"}), "CutCopyMode": NUMBER + ["False", "xlCopy", "xlCut"],
}
#: A Workbook property that outlasts its workbook, in the user's registry: put back like Application's.
USER_LEVEL = {"DoNotPromptForConvert"}
#: The worksheet's sort settings, each a choice of two or three values.
SORT = {
    "Header": NUMBER + ["xlYes", "xlNo", "xlGuess"], "MatchCase": BOOLEAN,
    "Orientation": NUMBER + ["xlTopToBottom", "xlLeftToRight"], "SortMethod": NUMBER + ["xlPinYin", "xlStroke"],
}
#: A sort field's settings, on the field Add makes on A1.
SORT_FIELD = {
    "Order": NUMBER + ["xlAscending", "xlDescending"], "DataOption": NUMBER + ["xlSortNormal", "xlSortTextAsNumbers"],
    "SubField": TEXT,
}
TARGETS = {"Worksheet": "ws", "Workbook": "wb", "Application": "Application", "Sort": "ws.Sort",
           "SortField": 'ws.Sort.SortFields.Add(ws.Range("A1"))'}
#: Where the typed half reaches the object, where that differs from how the object is made.
TYPED = {"SortField": "ws.Sort.SortFields(1)"}
OBJECTS = {"Worksheet": WORKSHEET, "Workbook": WORKBOOK, "Application": APPLICATION, "Sort": SORT, "SortField": SORT_FIELD}

TEMPLATE = '''Private Function Read1(o As Object) As String
    On Error GoTo Failed
    Read1 = TypeName(o.{prop}) & ":" & CStr(o.{prop})
    Exit Function
Failed:
    Read1 = "E" & Err.Number
End Function

Public Function T(ByVal which As Long) As String
    Dim wb As Workbook, ws As Worksheet, o As Object, s As String, was As Variant, hadOld As Boolean
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    wb.Worksheets.Add After:=ws
    ws.Activate
    ws.Range("A1").Value = 1
    Set o = {target}
    {keep}
    s = Read1(o)
    On Error Resume Next
    Err.Clear
    Select Case which
{cases}
    End Select
    s = s & "|E" & Err.Number & "|"
    On Error GoTo 0
    s = s & Read1(o)
    On Error Resume Next
    {restore}
    ws.Visible = True
    wb.Saved = True
    wb.IsAddin = False
    Application.DisplayAlerts = False
    wb.Close False
    T = s
End Function
'''

ALL = '''
Public Function RunAll(ByVal last As Long) As String
    Dim i As Long, s As String
    For i = 0 To last
        s = s & "@@@" & T(i)
    Next
    RunAll = Mid$(s, 4)
End Function
'''


def source(obj: str, prop: str, values: list[str]) -> str:
    late = [f"    Case {index}\n        o.{prop} = {value}" for index, value in enumerate(values)]
    typed = [f"    Case {index + len(values)}\n        {TYPED.get(obj, TARGETS[obj])}.{prop} = {value}" for index, value in enumerate(values)]
    cases = "\n".join([*late, *typed])
    keep = restore = ""
    if obj == "Application" or prop in USER_LEVEL:
        keep = f"was = {TARGETS[obj]}.{prop}"
        restore = f"{TARGETS[obj]}.{prop} = was"
    return TEMPLATE.format(prop=prop, target=TARGETS[obj], cases=cases, keep=keep, restore=restore)


def save(data: dict[str, dict[str, object]]) -> None:
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "rows": data}, indent=1) + "\n",
                   encoding="utf-8")


def measure(objects: list[str], data: dict[str, dict[str, object]]) -> None:
    """Saved after each property, and a property already measured is not measured again, so a run can be taken up. The
    harness lets one session run on a machine at a time, so each property takes its own and lets others in between."""
    for obj in objects:
        for prop, values in OBJECTS[obj].items():
            name = f"{obj}.{prop}"
            if name in data and not str(data[name]["answers"][0]).startswith("failed harness"):  # type: ignore[index]
                continue
            text = source(obj, prop, values)
            answers: list[str]
            try:
                with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                    excel.new_document()
                    result = excel.run_vba(text + ALL, "RunAll", (2 * len(values) - 1,), timeout=120.0)
                answers = str(result.value).split("@@@") if result.ok else [f"failed {result.outcome}"] * 2 * len(values)
            except HarnessError as failure:
                answers = [f"failed harness {str(failure)[:60]}"] * 2 * len(values)
            data[name] = {"object": obj, "property": prop, "values": values, "source": text, "answers": answers}
            save(data)
            print(obj, prop, answers[0][:70], flush=True)


def main() -> None:
    which = sys.argv[1] if len(sys.argv) > 1 else "books"
    data: dict[str, dict[str, object]] = {}
    if OUT.exists():
        data = json.loads(OUT.read_text(encoding="utf-8"))["rows"]
    measure({"application": ["Application"], "sort": ["Sort", "SortField"]}.get(which, ["Worksheet", "Workbook"]), data)


if __name__ == "__main__":
    main()
