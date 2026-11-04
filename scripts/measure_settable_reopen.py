"""What a setting of Worksheet, Workbook and Application is after a save and reopen, in live Excel, with what the sheet and
workbook parts hold of it. A new workbook takes each value, is saved to a temporary folder, closed and opened again.
Rules for EnableFormatConditionsCalculation are split as well: its reading lags a setting by some rule that adding a sheet
is the suspect for.

    python scripts/measure_settable_reopen.py

writes tests/fixtures/settable_reopen.json, which tests/test_excel_settable_reopen.py replays.
"""
from __future__ import annotations

import json
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/settable_reopen.json"

#: Each: target, member, a value other than the new workbook's.
SETTINGS = [
    ("ws", "DisplayRightToLeft", "True"), ("ws", "EnableCalculation", "False"), ("ws", "EnableOutlining", "True"),
    ("ws", "EnableAutoFilter", "True"), ("ws", "EnablePivotTable", "True"), ("ws", "EnableSelection", "1"),
    ("ws", "Name", '"Renamed"'), ("ws", "ScrollArea", '"A1:B2"'), ("ws", "StandardWidth", "12"),
    ("ws", "TransitionExpEval", "True"), ("ws", "TransitionFormEntry", "True"), ("other", "Visible", "0"),
    ("other", "Visible", "2"), ("ws", "DisplayPageBreaks", "True"),
    ("wb", "Date1904", "True"), ("wb", "ForceFullCalculation", "True"), ("wb", "PrecisionAsDisplayed", "True"),
    ("wb", "SaveLinkValues", "False"), ("wb", "ReadOnlyRecommended", "False"), ("wb", "RemovePersonalInformation", "True"),
    ("wb", "CheckCompatibility", "True"), ("wb", "DisplayDrawingObjects", "3"), ("wb", "DisplayDrawingObjects", "2"),
    ("wb", "UpdateRemoteReferences", "False"), ("wb", "ShowPivotTableFieldList", "False"),
    ("wb", "ShowPivotChartActiveFields", "True"), ("wb", "EnableAutoRecover", "False"), ("wb", "UpdateLinks", "2"),
    ("wb", "ChartDataPointTrack", "False"), ("wb", "DisplayInkComments", "False"), ("wb", "InactiveListBorderVisible", "False"),
    ("wb", "PersonalViewListSettings", "False"), ("wb", "PersonalViewPrintSettings", "False"),
    ("wb", "TemplateRemoveExtData", "True"), ("wb", "DefaultTableStyle", '"TableStyleLight1"'),
    ("wb", "DefaultPivotTableStyle", '"PivotStyleDark1"'), ("wb", "KeepChangeHistory", "False"),
    ("Application", "Calculation", "xlCalculationManual"), ("Application", "Calculation", "2"),
]
LAG = {
    "one_sheet": "ws.EnableFormatConditionsCalculation = False\ns = ws.EnableFormatConditionsCalculation",
    "added_first": 'wb.Worksheets.Add After:=ws\nws.Activate\nws.EnableFormatConditionsCalculation = False\ns = ws.EnableFormatConditionsCalculation',
    "added_after": 'ws.EnableFormatConditionsCalculation = False\ns = ws.EnableFormatConditionsCalculation\n'
                   'wb.Worksheets.Add After:=ws\ns = s & "," & ws.EnableFormatConditionsCalculation',
    "added_not_activated": 'wb.Worksheets.Add After:=ws\nws.EnableFormatConditionsCalculation = False\ns = ws.EnableFormatConditionsCalculation',
    "on_the_added": 'Dim n As Worksheet\nSet n = wb.Worksheets.Add(After:=ws)\nn.EnableFormatConditionsCalculation = False\ns = n.EnableFormatConditionsCalculation',
    "set_twice": "ws.EnableFormatConditionsCalculation = False\nws.EnableFormatConditionsCalculation = False\ns = ws.EnableFormatConditionsCalculation",
    "written_cell": 'ws.Range("A1").Value = 1\nws.EnableFormatConditionsCalculation = False\ns = ws.EnableFormatConditionsCalculation',
}

REOPEN = '''Function Reopen(ByVal folder As String) As String
    Dim wb As Workbook, ws As Worksheet, other As Worksheet, s As String, file As String, was As Variant
    Application.DisplayAlerts = False
    was = Application.Calculation
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    Set other = wb.Worksheets.Add(After:=ws)
    ws.Activate
    On Error GoTo Failed
    {target}.{member} = {value}
    s = "set=" & CStr({read})
    file = folder & "\\reopen.xlsx"
    wb.SaveAs file, 51
    wb.Close False
    Application.Calculation = xlCalculationAutomatic
    Set wb = Workbooks.Open(file)
    Set ws = wb.Worksheets(1): Set other = wb.Worksheets(2)
    s = s & "|reopened=" & CStr({read})
    GoTo Done
Failed:
    s = s & "|E" & Err.Number
Done:
    On Error Resume Next
    Application.Calculation = was
    wb.Close False
    Kill file
    Reopen = s
End Function
'''

LAG_PROBE = '''Function Lag() As String
    Dim wb As Workbook, ws As Worksheet, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    {action}
    wb.Close False
    Lag = s
End Function
'''


def reopen_source(target: str, member: str, value: str) -> str:
    read = f"{target}.{member}"
    return REOPEN.format(target=target, member=member, value=value, read=read)


def main() -> None:
    folder = tempfile.mkdtemp(prefix="pyopenvba_reopen_")
    settings: list[dict[str, str]] = []
    lag: dict[str, str] = {}
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
            excel.new_document()
            for name, action in LAG.items():
                result = excel.run_vba(LAG_PROBE.format(action=action), "Lag", timeout=60.0)
                lag[name] = str(result.value) if result.ok else f"failed {result.outcome}"
                print(name, lag[name], flush=True)
        for target, member, value in SETTINGS:
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                result = excel.run_vba(reopen_source(target, member, value), "Reopen", (folder,), timeout=60.0)
            answer = str(result.value) if result.ok else f"failed {result.outcome}"
            print(target, member, value, answer, flush=True)
            settings.append({"target": target, "member": member, "value": value, "answer": answer})
            OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "lag": lag, "lag_probe": LAG_PROBE,
                                       "lag_actions": LAG, "probe": REOPEN, "settings": settings}, indent=1) + "\n",
                           encoding="utf-8")
    finally:
        shutil.rmtree(folder, ignore_errors=True)


if __name__ == "__main__":
    main()
