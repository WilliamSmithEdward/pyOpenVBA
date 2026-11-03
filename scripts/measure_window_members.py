"""Window and Windows members that need no measure of the screen: ActivePane, walking Windows, EnableResize and
ScrollWorkbookTabs. Nothing here moves or sizes a window, which Excel would save for its user.

    python scripts/measure_window_members.py

writes tests/fixtures/window_members.json, which tests/test_excel_window_members.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/window_members.json"

ROWS = {
    "active_pane_plain": 'Note TypeName(ActiveWindow.ActivePane) & "," & ActiveWindow.ActivePane.Index',
    "active_pane_is_item": 'Note ActiveWindow.ActivePane.Index & "," & ActiveWindow.Panes.Count',
    "active_pane_set": 'ActiveWindow.ActivePane = ActiveWindow.Panes(1)',
    "walk_windows": 'Dim w As Window\nFor Each w In Windows\nNote w.Caption\nNext',
    "walk_two": 'Workbooks.Add\nDim w As Window\nFor Each w In Windows\nNote TypeName(w) & w.Index\nNext',
    "walk_book_windows": 'Dim w As Window\nFor Each w In ActiveWorkbook.Windows\nNote w.Index\nNext',
    "enable_resize_read": 'Note ActiveWindow.EnableResize',
    "scroll_tabs_position": 'ActiveWindow.ScrollWorkbookTabs Position:=xlFirst\nNote "ok"',
    "scroll_tabs_sheets": 'ActiveWindow.ScrollWorkbookTabs Sheets:=1\nNote "ok"',
    "scroll_tabs_none": 'ActiveWindow.ScrollWorkbookTabs\nNote "ok"',
    "scroll_tabs_both": 'ActiveWindow.ScrollWorkbookTabs Sheets:=1, Position:=xlFirst\nNote "ok"',
    "scroll_tabs_bad_position": 'ActiveWindow.ScrollWorkbookTabs Position:=7\nNote "ok"',
    "data_type_icons": 'Note ActiveWindow.DisplayDataTypeIcons',
    "arrange_count": 'Note Windows.Count',
    "sync_read": 'Note Windows.SyncScrollingSideBySide',
    "break_side_none": 'Note Windows.BreakSideBySide',
    "reset_side_none": 'Windows.ResetPositionsSideBySide\nNote "ok"',
    "compare_self": 'Note Windows.CompareSideBySideWith(ActiveWindow.Caption)',
    "compare_nothing": 'Note Windows.CompareSideBySideWith("Nowhere")',
}

PROBE = """Dim Log As String
Sub Note(ByVal s As String)
Log = Log & s & "|"
End Sub
Function Probe() As String
Application.DisplayAlerts = False
Dim book As Workbook, home As Workbook
Set home = ActiveWorkbook
On Error GoTo Failed
{action}
GoTo Done
Failed:
Log = Log & "!E" & Err.Number
Done:
On Error Resume Next
For Each book In Workbooks
If book.Name <> home.Name Then book.Close False
Next
Probe = Log
End Function
"""


def main() -> None:
    rows: list[dict[str, str]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=60.0)) as excel:
        for name, action in ROWS.items():
            excel.new_document()
            result = excel.run_vba(PROBE.format(action=action), "Probe", timeout=30.0)
            value = str(result.value) if result.ok else f"failed {result.outcome} {result.message}"
            print(name, value, flush=True)
            rows.append({"name": name, "action": action, "result": value})
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "rows": rows},
                              indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
