"""Application.OnTime in live Excel: when a procedure runs, in what order, how Schedule:=False finds it, and what Excel
refuses. Each row is a macro that schedules, then a wait outside any macro, then a read of the log; Excel runs a due
procedure only when no macro is running, unless the row says otherwise.

    python scripts/measure_on_time.py

writes tests/fixtures/on_time.json, which tests/test_excel_on_time.py replays.
"""
from __future__ import annotations

import json
import time
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/on_time.json"

#: The log is a cell: the harness replaces its own module between calls, which resets every module's variables.
TIMERS = """Public Sub Note(ByVal s As String)
    ThisWorkbook.Worksheets(1).Range("Z1").Value = "'" & Logged() & s & "|"
End Sub
Public Function Logged() As String
    Logged = ThisWorkbook.Worksheets(1).Range("Z1").Value
End Function
Public Sub A()
    Note "A"
End Sub
Public Sub B()
    Note "B"
End Sub
Public Sub WithArgument(ByVal x As Long)
    Note "arg" & x
End Sub
Public Sub Again()
    Note "again"
    Application.OnTime Now, "A"
End Sub
Private Sub Hidden()
    Note "hidden"
End Sub
"""

#: Each row: what the macro does; its answer is what it logged by the time it ended and any error, and the log is read
#: again after the macro has ended and Excel has had time to run what is due.
ROWS = {
    "due_now": 'Application.OnTime Now, "A"',
    "in_the_past": 'Application.OnTime Now - 1, "A"',
    "in_the_future": 'Dim t As Date\nt = Now + 1\nApplication.OnTime t, "A"\nNote "set"\nApplication.OnTime t, "A", , False',
    "not_during_wait": 'Application.OnTime Now, "A"\nApplication.Wait Now + TimeSerial(0, 0, 1)',
    "not_during_do_events": 'Application.OnTime Now, "A"\nDim i As Long\nFor i = 1 To 2000\nDoEvents\nNext',
    "twice_same": 'Dim t As Date\nt = Now\nApplication.OnTime t, "A"\nApplication.OnTime t, "A"',
    "two_same_time": 'Dim t As Date\nt = Now\nApplication.OnTime t, "B"\nApplication.OnTime t, "A"',
    "two_by_time": 'Dim t As Date\nt = Now\nApplication.OnTime t - 1 / 86400, "B"\nApplication.OnTime t - 2 / 86400, "A"',
    "two_by_time_other_way": 'Dim t As Date\nt = Now\nApplication.OnTime t - 2 / 86400, "B"\nApplication.OnTime t - 1 / 86400, "A"',
    "cancel": 'Dim t As Date\nt = Now + 1\nApplication.OnTime t, "A"\nApplication.OnTime t, "A", , False',
    "cancel_due": 'Dim t As Date\nt = Now\nApplication.OnTime t, "A"\nApplication.OnTime t, "A", , False',
    "cancel_none": 'Application.OnTime Now + 1, "A", , False',
    "cancel_other_time": 'Dim t As Date\nt = Now + 1\nApplication.OnTime t, "A"\nOn Error Resume Next\n'
                         'Application.OnTime t + 1, "A", , False\nNote "E" & Err.Number\nOn Error GoTo 0\n'
                         'Application.OnTime t, "A", , False',
    "cancel_other_name": 'Dim t As Date\nt = Now + 1\nApplication.OnTime t, "A"\nOn Error Resume Next\n'
                         'Application.OnTime t, "B", , False\nNote "E" & Err.Number\nOn Error GoTo 0\n'
                         'Application.OnTime t, "A", , False',
    "cancel_other_case": 'Dim t As Date\nt = Now + 1\nApplication.OnTime t, "A"\nOn Error Resume Next\n'
                         'Application.OnTime t, "a", , False\nNote "E" & Err.Number\nOn Error GoTo 0\n'
                         'If Err.Number = 0 Then Note "ok"\nOn Error Resume Next\n'
                         'Application.OnTime t, "A", , False',
    "cancel_twice": 'Dim t As Date\nt = Now + 1\nApplication.OnTime t, "A"\nApplication.OnTime t, "A", , False\n'
                    'Application.OnTime t, "A", , False',
    "cancel_one_of_two": 'Dim t As Date\nt = Now\nApplication.OnTime t, "A"\nApplication.OnTime t, "A"\n'
                         'Application.OnTime t, "A", , False',
    "schedule_true": 'Application.OnTime Now, "A", , True',
    "schedule_text": 'Application.OnTime Now, "A", , "True"',
    "latest_past": 'Application.OnTime Now - 2, "A", Now - 1',
    "latest_future": 'Application.OnTime Now - 2, "A", Now + 1',
    "time_of_day_past": 'Application.OnTime TimeValue("00:00:00"), "A", , True\n'
                        'Application.OnTime TimeValue("00:00:00"), "A", , False',
    "number": 'Application.OnTime CDbl(Now), "A"',
    "text_time": 'Application.OnTime CStr(Now), "A"',
    "bad_time": 'Application.OnTime "soon", "A"',
    "null_time": 'Application.OnTime Null, "A"',
    "empty_name": 'Application.OnTime Now, ""',
    "qualified": 'Application.OnTime Now, "Timers.A"',
    "booked": 'Application.OnTime Now, "\'" & ActiveWorkbook.Name & "\'!A"',
    "argument": 'Application.OnTime Now, "\'WithArgument 7\'"',
    "private": 'Application.OnTime Now, "Timers.Hidden"',
    "from_timer": 'Application.OnTime Now, "Again"',
    "do_events_two_calls": 'Application.OnTime Now - 1, "A"\nDoEvents\nDoEvents',
    "do_events_two": 'Application.OnTime Now, "A"\nApplication.OnTime Now, "B"\nDim i As Long\nFor i = 1 To 2000\nDoEvents\nNext',
    "do_events_once": 'Application.OnTime Now - 1, "A"\nDoEvents',
    "do_events_future": 'Dim t As Date\nt = Now + 1\nApplication.OnTime t, "A"\nDoEvents\nApplication.OnTime t, "A", , False',
}

PROBE = """Function Probe() As String
On Error GoTo Failed
{action}
Probe = Logged()
Exit Function
Failed:
Probe = Logged() & "!E" & Err.Number
End Function
"""

READ = 'Function ReadLog() As String\nReadLog = Logged()\nEnd Function\n'


def main() -> None:
    rows: list[dict[str, str]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        for name, action in ROWS.items():
            excel.new_document()
            excel.add_module("Timers", TIMERS)
            during = excel.run_vba(PROBE.format(action=action), "Probe", timeout=30.0)
            time.sleep(1.5)
            after = excel.run_vba(READ, "ReadLog", timeout=30.0)
            row = {"name": name, "action": action,
                   "during": str(during.value) if during.ok else f"failed {during.outcome} {during.message}",
                   "after": str(after.value) if after.ok else f"failed {after.outcome} {after.message}"}
            print(name, row["during"], "=>", row["after"], flush=True)
            rows.append(row)
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "timers": TIMERS, "probe": PROBE,
                               "rows": rows}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
